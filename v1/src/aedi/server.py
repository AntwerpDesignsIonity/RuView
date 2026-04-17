#!/usr/bin/env python3
"""
AEDI — Automated Ecosystem Designs Ionity
AI Assistant Service for AEDI-S WiFi Sensing Platform.

Runs a lightweight HTTP API on port 3002 that:
  - Connects to Ollama (Gemma 4 E2B) for local LLM inference
  - Pulls live context from the sensing-server (port 3000)
  - Serves as the brain behind the AEDI chat UI

By Antwerp Designs — Johan Wilhelm van Antwerp
Ionity (Pty) Ltd — www.ionity.today
"""

import asyncio
import json
import logging
import os
import signal
import sys
import time
import traceback
from datetime import datetime, timezone
from urllib.parse import urlparse, parse_qs
from urllib.request import urlopen, Request
from urllib.error import URLError

import aiohttp
from aiohttp import web

# ─── Config ────────────────────────────────────────────────────────────────────
AEDI_PORT      = int(os.environ.get("AEDI_PORT", 3002))
OLLAMA_URL     = os.environ.get("OLLAMA_URL", "http://localhost:11434")
OLLAMA_MODEL   = os.environ.get("OLLAMA_MODEL", "gemma4:e2b")
SENSING_URL    = os.environ.get("SENSING_URL", "http://localhost:3000")
AEDI_BIND      = os.environ.get("AEDI_BIND", "0.0.0.0")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [AEDI] %(levelname)s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("aedi")

# ─── AEDI System Prompt ───────────────────────────────────────────────────────
SYSTEM_PROMPT = """You are AEDI — Automated Ecosystem Designs Ionity.

Your identity:
• You are the AI assistant for Ionity (Pty) Ltd, an AioT (AI + IoT) company.
• Created by Antwerp Designs, founded by Johan Wilhelm van Antwerp.
• Based in Centurion/Pretoria, South Africa.
• Website: www.ionity.today | Email: ai@ionity.today
• Always refer to yourself as AEDI when asked who you are.

About Ionity:
• Ionity builds automated, energy-efficient ecosystems connecting edge hardware with Cloud AI.
• Core stack: IoT Edge Sensors → Cloud Infrastructure → AI-Driven Automation.
• Products: Ai-M Motherboard (24GB RAM edge AI), Edge Computing System, AEDI Cloud Command & Digital Twin Center, AEDI Neural Inference Engine & API.
• Services: Cloud AI & Predictive Analytics, Edge-to-Cloud Middleware, Zero-Trust IoT Security, RTLS & Proximity Services, Data Services & Cloud Storage, AIaaS & MLOps.
• Industries: Transport, Energy, Medical, Banking, Security, GIS, Mining, Hospitality, Retail, Corporate IT, Industrial/Manufacturing.
• Specialties: Predictive Maintenance (PdM), Data Acquisition (DAQ), Digital Twins, SCADA retrofitting, OTA firmware updates, self-healing mesh networks.

About AEDI-S (the system you're running on):
• AEDI-S is Ionity's WiFi-based human sensing & pose estimation platform.
• Uses WiFi Channel State Information (CSI) to detect presence, motion, breathing, and body pose through walls — no cameras needed.
• Hardware: Raspberry Pi 5 hub + ESP32-S3 nodes forming a multistatic WiFi sensing mesh.
• Stack: Rust sensing-server (Axum) + Python ML pipeline + 3D Observatory visualization.
• Crates: 15 Rust crates including signal processing (RuvSense), neural networks, training, MAT (Mass Casualty Assessment), hardware drivers, WASM bindings.
• RuvSense modules: multiband fusion, phase alignment, multistatic sensing, coherence gating, pose tracking (17-keypoint Kalman), RF tomography, gesture recognition, adversarial detection.
• Current deployment: 3 ESP32-S3 nodes on "Ionity 2.4" WiFi network.

Your capabilities:
• Analyze live WiFi sensing data (RSSI, CSI, motion detection, node health).
• Explain system architecture, signal processing, and sensing concepts.
• Give tips on improving sensing coverage, node placement, and signal quality.
• Help troubleshoot hardware issues (ESP32, WiFi connectivity, UDP CSI).
• Suggest improvements to the codebase and deployment.
• Answer questions about Ionity products and services.

Guidelines:
• Be technical but approachable — you serve engineers and decision-makers.
• When you have live sensing data context, reference it in your responses.
• Be concise and direct — users are hands-on builders.
• For Ionity info, reference www.ionity.today.
• Use your knowledge to provide actionable insights and real-time analysis.
"""

# ─── Shared state ──────────────────────────────────────────────────────────────
conversations: dict[str, list] = {}        # session_id → messages
_sensing_cache: dict = {}
_cache_ts: float = 0.0

# ─── Helpers ───────────────────────────────────────────────────────────────────
async def fetch_sensing_context(session: aiohttp.ClientSession) -> dict:
    """Pull live data from sensing-server, cache for 2s."""
    global _sensing_cache, _cache_ts
    now = time.monotonic()
    if now - _cache_ts < 2.0 and _sensing_cache:
        return _sensing_cache
    try:
        endpoints = {
            "sensing": f"{SENSING_URL}/api/sensing",
            "health": f"{SENSING_URL}/health",
            "autorepair": f"{SENSING_URL}/health/autorepair",
        }
        ctx = {}
        for key, url in endpoints.items():
            try:
                async with session.get(url, timeout=aiohttp.ClientTimeout(total=2)) as r:
                    if r.status == 200:
                        ctx[key] = await r.json()
            except Exception:
                pass
        _sensing_cache = ctx
        _cache_ts = now
        return ctx
    except Exception:
        return _sensing_cache


def build_context_summary(ctx: dict) -> str:
    """Format sensing data into a readable context string for the LLM."""
    parts = []
    if "sensing" in ctx:
        s = ctx["sensing"]
        parts.append(
            f"[LIVE SENSING] Source: {s.get('source','?')}, Tick: {s.get('tick',0)}, "
            f"Nodes: {len(s.get('nodes',[]))}, Motion: {s.get('motion','?')}, "
            f"Persons: {len(s.get('persons',[]))}"
        )
        for node in s.get("nodes", []):
            parts.append(
                f"  Node {node.get('bssid','?')}: RSSI={node.get('rssi','?')} dBm, "
                f"Channel={node.get('channel','?')}, Packets={node.get('packets','?')}"
            )
    if "health" in ctx:
        h = ctx["health"]
        parts.append(
            f"[SYSTEM HEALTH] Status: {h.get('status','?')}, Clients: {h.get('clients',0)}, "
            f"Source: {h.get('source','?')}"
        )
    if "autorepair" in ctx:
        a = ctx["autorepair"]
        # Handle nested format: {"autorepair": {...}} or flat
        if "autorepair" in a:
            a = a["autorepair"]
        parts.append(
            f"[AUTOREPAIR] Healthy: {a.get('healthy','?')}, Memory: {a.get('memory_mb', a.get('rss_mb','?'))} MB, "
            f"Uptime: {a.get('uptime_secs','?')}s, Source: {a.get('source','?')}, "
            f"Tick: {a.get('tick','?')}"
        )
    return "\n".join(parts) if parts else "[No live sensing data available]"


async def ollama_chat(session: aiohttp.ClientSession, messages: list, stream: bool = False):
    """Send chat request to Ollama. Returns full response or async generator for streaming."""
    payload = {
        "model": OLLAMA_MODEL,
        "messages": messages,
        "stream": stream,
        "options": {
            "temperature": 1.0,
            "top_p": 0.95,
            "top_k": 64,
            "num_predict": 256,  # Keep short for RPi 5 CPU-only
        },
    }
    url = f"{OLLAMA_URL}/api/chat"
    if stream:
        resp = await session.post(url, json=payload, timeout=aiohttp.ClientTimeout(total=600))
        return resp  # caller reads the stream
    else:
        async with session.post(url, json=payload, timeout=aiohttp.ClientTimeout(total=600)) as resp:
            if resp.status != 200:
                text = await resp.text()
                return {"error": f"Ollama returned {resp.status}: {text}"}
            return await resp.json()


async def check_ollama(session: aiohttp.ClientSession) -> dict:
    """Check Ollama status and available models."""
    try:
        async with session.get(f"{OLLAMA_URL}/api/tags", timeout=aiohttp.ClientTimeout(total=5)) as r:
            if r.status == 200:
                data = await r.json()
                models = [m.get("name", "?") for m in data.get("models", [])]
                return {"online": True, "models": models, "target": OLLAMA_MODEL}
            return {"online": False, "error": f"Status {r.status}"}
    except Exception as e:
        return {"online": False, "error": str(e)}


# ─── HTTP Handlers ─────────────────────────────────────────────────────────────

async def handle_health(request: web.Request) -> web.Response:
    """GET /health — AEDI service health."""
    session = request.app["http_session"]
    ollama = await check_ollama(session)
    sensing_ctx = await fetch_sensing_context(session)
    return web.json_response({
        "service": "aedi",
        "status": "ok",
        "model": OLLAMA_MODEL,
        "ollama": ollama,
        "sensing_connected": bool(sensing_ctx),
        "active_conversations": len(conversations),
        "timestamp": datetime.now(timezone.utc).isoformat(),
    })


async def handle_chat(request: web.Request) -> web.Response:
    """POST /api/chat — Send a message, get AEDI's response."""
    session = request.app["http_session"]
    try:
        body = await request.json()
    except Exception:
        return web.json_response({"error": "Invalid JSON"}, status=400)

    user_msg = body.get("message", "").strip()
    if not user_msg:
        return web.json_response({"error": "Empty message"}, status=400)
    if len(user_msg) > 4000:
        return web.json_response({"error": "Message too long (max 4000 chars)"}, status=400)

    session_id = body.get("session_id", "default")
    # Sanitize session_id
    session_id = "".join(c for c in session_id if c.isalnum() or c in "-_")[:64] or "default"

    # Get or create conversation
    if session_id not in conversations:
        conversations[session_id] = []
    conv = conversations[session_id]

    # Fetch live sensing context
    sensing_ctx = await fetch_sensing_context(session)
    context_str = build_context_summary(sensing_ctx)

    # Build messages for Ollama
    system_msg = SYSTEM_PROMPT
    if context_str:
        system_msg += f"\n\n--- CURRENT LIVE DATA ---\n{context_str}\n--- END LIVE DATA ---"

    messages = [{"role": "system", "content": system_msg}]

    # Keep last 20 messages for context window
    for msg in conv[-20:]:
        messages.append(msg)

    messages.append({"role": "user", "content": user_msg})

    # Call Ollama
    result = await ollama_chat(session, messages)
    if "error" in result:
        return web.json_response({"error": result["error"]}, status=502)

    assistant_msg = result.get("message", {}).get("content", "").strip()
    if not assistant_msg:
        assistant_msg = "I apologize, I was unable to generate a response. Please try again."

    # Store conversation
    conv.append({"role": "user", "content": user_msg})
    conv.append({"role": "assistant", "content": assistant_msg})

    # Trim conversation history
    if len(conv) > 40:
        conv[:] = conv[-40:]

    return web.json_response({
        "response": assistant_msg,
        "session_id": session_id,
        "model": result.get("model", OLLAMA_MODEL),
        "context": {
            "source": sensing_ctx.get("sensing", {}).get("source", "unknown"),
            "tick": sensing_ctx.get("sensing", {}).get("tick", 0),
            "healthy": sensing_ctx.get("health", {}).get("status") == "ok",
        },
        "eval_count": result.get("eval_count", 0),
        "eval_duration_ms": round(result.get("eval_duration", 0) / 1e6, 1),
    })


async def handle_chat_stream(request: web.Request) -> web.StreamResponse:
    """POST /api/chat/stream — SSE streaming chat."""
    session = request.app["http_session"]
    try:
        body = await request.json()
    except Exception:
        return web.json_response({"error": "Invalid JSON"}, status=400)

    user_msg = body.get("message", "").strip()
    if not user_msg:
        return web.json_response({"error": "Empty message"}, status=400)
    if len(user_msg) > 4000:
        return web.json_response({"error": "Message too long"}, status=400)

    session_id = body.get("session_id", "default")
    session_id = "".join(c for c in session_id if c.isalnum() or c in "-_")[:64] or "default"

    if session_id not in conversations:
        conversations[session_id] = []
    conv = conversations[session_id]

    sensing_ctx = await fetch_sensing_context(session)
    context_str = build_context_summary(sensing_ctx)

    system_msg = SYSTEM_PROMPT
    if context_str:
        system_msg += f"\n\n--- CURRENT LIVE DATA ---\n{context_str}\n--- END LIVE DATA ---"

    messages = [{"role": "system", "content": system_msg}]
    for msg in conv[-20:]:
        messages.append(msg)
    messages.append({"role": "user", "content": user_msg})

    # Start SSE response
    resp = web.StreamResponse(
        status=200,
        reason="OK",
        headers={
            "Content-Type": "text/event-stream",
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "Access-Control-Allow-Origin": "*",
        },
    )
    await resp.prepare(request)

    full_response = []
    client_gone = False
    try:
        ollama_resp = await ollama_chat(session, messages, stream=True)
        async for line in ollama_resp.content:
            if not line:
                continue
            try:
                chunk = json.loads(line)
                token = chunk.get("message", {}).get("content", "")
                if token:
                    full_response.append(token)
                    if not client_gone:
                        try:
                            event = json.dumps({"token": token})
                            await resp.write(f"data: {event}\n\n".encode())
                        except (ConnectionResetError, ConnectionError, Exception) as we:
                            if "closing transport" in str(we) or "ConnectionReset" in type(we).__name__:
                                client_gone = True
                                log.info("Client disconnected, continuing inference to save conversation")
                            else:
                                raise
                if chunk.get("done"):
                    if not client_gone:
                        try:
                            done_event = json.dumps({
                                "done": True,
                                "model": chunk.get("model", OLLAMA_MODEL),
                                "eval_count": chunk.get("eval_count", 0),
                            })
                            await resp.write(f"data: {done_event}\n\n".encode())
                        except (ConnectionResetError, ConnectionError, Exception):
                            client_gone = True
            except json.JSONDecodeError:
                pass
    except (ConnectionResetError, ConnectionError) as e:
        if "closing transport" in str(e) or "ConnectionReset" in type(e).__name__:
            client_gone = True
        else:
            log.error(f"Stream error: {e}")
    except asyncio.CancelledError:
        client_gone = True
    except Exception as e:
        if not client_gone:
            try:
                err = json.dumps({"error": str(e)})
                await resp.write(f"data: {err}\n\n".encode())
            except Exception:
                pass
        log.error(f"Stream error: {e}")

    # Store conversation even if client disconnected
    assistant_text = "".join(full_response)
    if assistant_text:
        conv.append({"role": "user", "content": user_msg})
        conv.append({"role": "assistant", "content": assistant_text})
        if len(conv) > 40:
            conv[:] = conv[-40:]

    if not client_gone:
        try:
            await resp.write_eof()
        except Exception:
            pass
    return resp


async def handle_context(request: web.Request) -> web.Response:
    """GET /api/context — Current live sensing context."""
    session = request.app["http_session"]
    ctx = await fetch_sensing_context(session)
    return web.json_response({
        "raw": ctx,
        "summary": build_context_summary(ctx),
        "timestamp": datetime.now(timezone.utc).isoformat(),
    })


async def handle_conversations(request: web.Request) -> web.Response:
    """GET /api/conversations — List active sessions."""
    return web.json_response({
        "sessions": {
            sid: {"message_count": len(msgs), "last": msgs[-1]["content"][:80] if msgs else ""}
            for sid, msgs in conversations.items()
        }
    })


async def handle_clear(request: web.Request) -> web.Response:
    """POST /api/clear — Clear a conversation."""
    try:
        body = await request.json()
    except Exception:
        body = {}
    sid = body.get("session_id", "default")
    if sid in conversations:
        del conversations[sid]
    return web.json_response({"cleared": sid})


async def handle_cors_preflight(request: web.Request) -> web.Response:
    """Handle CORS preflight for cross-origin requests from the sensing-server UI."""
    return web.Response(
        status=200,
        headers={
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
            "Access-Control-Max-Age": "3600",
        },
    )


@web.middleware
async def cors_middleware(request, handler):
    """Add CORS headers to all responses."""
    if request.method == "OPTIONS":
        return await handle_cors_preflight(request)
    resp = await handler(request)
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


# ─── App Setup ─────────────────────────────────────────────────────────────────

async def on_startup(app: web.Application):
    app["http_session"] = aiohttp.ClientSession()
    log.info("AEDI service starting on %s:%d", AEDI_BIND, AEDI_PORT)
    log.info("Ollama: %s  Model: %s", OLLAMA_URL, OLLAMA_MODEL)
    log.info("Sensing: %s", SENSING_URL)

    # Check Ollama on startup
    ollama = await check_ollama(app["http_session"])
    if ollama.get("online"):
        log.info("Ollama online — models: %s", ", ".join(ollama["models"]))
        if not any(OLLAMA_MODEL.split(":")[0] in m for m in ollama["models"]):
            log.warning("Model '%s' not found. Pull it with: ollama pull %s", OLLAMA_MODEL, OLLAMA_MODEL)
    else:
        log.warning("Ollama offline (%s). Start Ollama first.", ollama.get("error", "?"))


async def on_shutdown(app: web.Application):
    await app["http_session"].close()
    log.info("AEDI service stopped.")


def create_app() -> web.Application:
    app = web.Application(middlewares=[cors_middleware])
    app.on_startup.append(on_startup)
    app.on_shutdown.append(on_shutdown)

    # Routes
    app.router.add_get("/health", handle_health)
    app.router.add_post("/api/chat", handle_chat)
    app.router.add_post("/api/chat/stream", handle_chat_stream)
    app.router.add_get("/api/context", handle_context)
    app.router.add_get("/api/conversations", handle_conversations)
    app.router.add_post("/api/clear", handle_clear)

    return app


def main():
    app = create_app()
    log.info("Starting AEDI — Automated Ecosystem Designs Ionity")
    web.run_app(app, host=AEDI_BIND, port=AEDI_PORT, print=None)


if __name__ == "__main__":
    main()
