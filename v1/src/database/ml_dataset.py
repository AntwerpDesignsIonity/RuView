"""
PyTorch Dataset that reads directly from the ruview_sensing.sqlite3 database.

Designed for RPi-friendly training: reads windows of CSI features from SQLite,
returns (X, y) pairs for presence/motion classification or vital-sign regression.

Usage::

    from v1.src.database.ml_dataset import SensingDataset
    from torch.utils.data import DataLoader

    ds = SensingDataset(window_size=20, target="presence")
    loader = DataLoader(ds, batch_size=32, shuffle=True)

    for X, y in loader:
        # X: (batch, window_size, n_features)
        # y: (batch,) int labels or (batch,) float regression targets
        pass
"""

from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import List, Optional, Tuple

import numpy as np

try:
    import torch
    from torch.utils.data import Dataset
except ImportError:
    raise ImportError("PyTorch is required: pip install torch")

_DEFAULT_DB = Path(__file__).resolve().parents[3] / "data" / "db" / "ruview_sensing.sqlite3"

# Feature columns for the model input (all numeric, from the readings table)
DEFAULT_FEATURES = [
    "feat_mean", "feat_std", "feat_variance", "feat_range",
    "feat_iqr", "feat_skewness", "feat_kurtosis",
    "feat_motion_power", "feat_breathing_power",
    "feat_dominant_freq_hz", "feat_spectral_power",
    "rssi_dbm", "mean_amplitude", "signal_quality",
    "esp_temp_c", "esp_wifi_rssi",
]


class SensingDataset(Dataset):
    """
    Sliding-window PyTorch dataset backed by SQLite.

    Parameters
    ----------
    db_path : path to ruview_sensing.sqlite3
    window_size : number of consecutive ticks per sample
    step : stride between windows (default: window_size // 2)
    features : list of column names from the readings table
    target : "presence" (binary), "motion" (4-class), or "breathing_bpm" (regression)
    session_id : restrict to one session (None = all data)
    device_id : restrict to one device
    min_quality : skip windows with avg signal_quality below this
    """

    MOTION_MAP = {"absent": 0, "idle": 1, "active": 2, "rapid": 3}

    def __init__(
        self,
        db_path: Optional[str] = None,
        window_size: int = 20,
        step: Optional[int] = None,
        features: Optional[List[str]] = None,
        target: str = "presence",
        session_id: Optional[str] = None,
        device_id: Optional[str] = None,
        min_quality: float = 0.0,
    ) -> None:
        self.db_path = Path(db_path) if db_path else _DEFAULT_DB
        self.window_size = window_size
        self.step = step or max(1, window_size // 2)
        self.features = features or DEFAULT_FEATURES
        self.target = target
        self.min_quality = min_quality

        # Load all qualifying rows into memory (efficient for RPi with ≤1M rows)
        conn = sqlite3.connect(str(self.db_path), timeout=10)
        conn.row_factory = sqlite3.Row

        cols = ", ".join(self.features + ["presence", "motion_level", "breathing_bpm", "signal_quality"])
        conditions = ["1=1"]
        params: list = []
        if session_id:
            conditions.append("session_id = ?")
            params.append(session_id)
        if device_id:
            conditions.append("device_id = ?")
            params.append(device_id)
        where = " AND ".join(conditions)

        rows = conn.execute(
            f"SELECT {cols} FROM readings WHERE {where} ORDER BY ts_epoch ASC",
            params,
        ).fetchall()
        conn.close()

        # Build numpy arrays
        feat_data = []
        target_data = []
        for r in rows:
            row_feats = [float(r[f]) if r[f] is not None else 0.0 for f in self.features]
            feat_data.append(row_feats)

            if target == "presence":
                target_data.append(int(r["presence"] or 0))
            elif target == "motion":
                target_data.append(self.MOTION_MAP.get(r["motion_level"], 0))
            elif target == "breathing_bpm":
                target_data.append(float(r["breathing_bpm"] or 0.0))
            else:
                target_data.append(0)

        self._feat_array = np.array(feat_data, dtype=np.float32)
        self._target_array = np.array(target_data, dtype=np.float32 if target == "breathing_bpm" else np.int64)

        # Build window indices
        self._indices: List[int] = []
        for i in range(0, len(self._feat_array) - window_size, self.step):
            # Optional quality filter
            if min_quality > 0:
                quality_col_idx = self.features.index("signal_quality") if "signal_quality" in self.features else -1
                if quality_col_idx >= 0:
                    avg_q = np.mean(self._feat_array[i:i + window_size, quality_col_idx])
                    if avg_q < min_quality:
                        continue
            self._indices.append(i)

    def __len__(self) -> int:
        return len(self._indices)

    def __getitem__(self, idx: int) -> Tuple[torch.Tensor, torch.Tensor]:
        start = self._indices[idx]
        X = torch.from_numpy(self._feat_array[start:start + self.window_size].copy())

        # Label: majority vote for classification, mean for regression
        targets_window = self._target_array[start:start + self.window_size]
        if self.target == "breathing_bpm":
            y = torch.tensor(float(np.mean(targets_window)), dtype=torch.float32)
        else:
            # Majority vote
            vals, counts = np.unique(targets_window, return_counts=True)
            y = torch.tensor(int(vals[np.argmax(counts)]), dtype=torch.long)

        return X, y

    @property
    def n_features(self) -> int:
        return len(self.features)

    @property
    def n_classes(self) -> int:
        if self.target == "presence":
            return 2
        elif self.target == "motion":
            return 4
        return 0  # regression
