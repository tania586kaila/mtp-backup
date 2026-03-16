"""
Anthropic-to-DeepSeek 代理服务
================================
将 Claude Code 发出的 Anthropic API 格式请求，
转换为 DeepSeek API 格式并转发，实现 Claude Code 使用 DeepSeek 模型。

原理：
  Claude Code 通过 ANTHROPIC_BASE_URL 环境变量指向本代理（http://127.0.0.1:8742）
  代理接收 /v1/messages 请求，转换格式后转发到 DeepSeek API

支持：
  - 普通请求（non-streaming）
  - 流式请求（streaming，SSE 格式）
  - 多轮对话（messages 数组）
  - system prompt

用法：
  1. 设置 DEEPSEEK_API_KEY 环境变量
  2. python deepseek_proxy.py
  3. 设置 ANTHROPIC_BASE_URL=http://127.0.0.1:8742
  4. 启动 Claude Code
"""

import json
import os
import sys
import logging
import requests
from flask import Flask, request, Response, jsonify

# ─────────────────────────────────────────────
# 配置
# ─────────────────────────────────────────────
PROXY_HOST = "127.0.0.1"
PROXY_PORT = 8742

# DeepSeek API 配置
DEEPSEEK_API_BASE = "https://api.deepseek.com"
DEEPSEEK_API_KEY  = os.environ.get("DEEPSEEK_API_KEY", "")

# 本机代理配置（解决地区限制）
# 自动读取系统代理，也可手动指定，如 "http://127.0.0.1:7897"
_sys_proxy = os.environ.get("HTTP_PROXY") or os.environ.get("http_proxy") or "http://127.0.0.1:7897"
REQUESTS_PROXIES = {
    "http":  _sys_proxy,
    "https": _sys_proxy,
}

# Claude 模型名 -> DeepSeek 模型名 映射
MODEL_MAP = {
    # Claude 3.x 系列 -> deepseek-chat（V3）
    "claude-opus-4-5":          "deepseek-chat",
    "claude-sonnet-4-5":        "deepseek-chat",
    "claude-haiku-3-5":         "deepseek-chat",
    "claude-3-5-sonnet-20241022": "deepseek-chat",
    "claude-3-5-haiku-20241022":  "deepseek-chat",
    "claude-3-opus-20240229":     "deepseek-chat",
    "claude-3-sonnet-20240229":   "deepseek-chat",
    "claude-3-haiku-20240307":    "deepseek-chat",
    # 别名
    "opus":   "deepseek-chat",
    "sonnet": "deepseek-chat",
    "haiku":  "deepseek-chat",
    # 如需使用推理模型，可改为 deepseek-reasoner
}
DEFAULT_DEEPSEEK_MODEL = "deepseek-chat"
# ─────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s - %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("ds_proxy")

app = Flask(__name__)


def map_model(claude_model: str) -> str:
    """将 Claude 模型名映射到 DeepSeek 模型名"""
    for key, val in MODEL_MAP.items():
        if key in claude_model.lower():
            return val
    return DEFAULT_DEEPSEEK_MODEL


def anthropic_to_openai_messages(anthropic_body: dict) -> list:
    """
    将 Anthropic messages 格式转换为 OpenAI/DeepSeek messages 格式。

    Anthropic 格式：
      messages: [{"role": "user", "content": [{"type": "text", "text": "..."}]}]

    OpenAI 格式：
      messages: [{"role": "user", "content": "..."}]
    """
    messages = []

    # system prompt 作为第一条 system 消息
    system = anthropic_body.get("system", "")
    if system:
        if isinstance(system, list):
            # system 可能是 content block 数组
            system_text = " ".join(
                block.get("text", "") for block in system
                if block.get("type") == "text"
            )
        else:
            system_text = str(system)
        if system_text.strip():
            messages.append({"role": "system", "content": system_text})

    # 转换 messages
    for msg in anthropic_body.get("messages", []):
        role = msg.get("role", "user")
        content = msg.get("content", "")

        if isinstance(content, str):
            text = content
        elif isinstance(content, list):
            # content block 数组，提取所有 text 类型
            text = "\n".join(
                block.get("text", "") for block in content
                if block.get("type") == "text"
            )
        else:
            text = str(content)

        messages.append({"role": role, "content": text})

    return messages


def openai_to_anthropic_response(openai_resp: dict, model: str) -> dict:
    """
    将 DeepSeek（OpenAI 格式）响应转换为 Anthropic 格式。
    """
    choice = openai_resp.get("choices", [{}])[0]
    message = choice.get("message", {})
    content_text = message.get("content", "")
    finish_reason = choice.get("finish_reason", "end_turn")

    # finish_reason 映射
    stop_reason_map = {
        "stop":         "end_turn",
        "length":       "max_tokens",
        "tool_calls":   "tool_use",
        "content_filter": "stop_sequence",
    }
    stop_reason = stop_reason_map.get(finish_reason, "end_turn")

    usage = openai_resp.get("usage", {})

    return {
        "id": openai_resp.get("id", "msg_proxy"),
        "type": "message",
        "role": "assistant",
        "model": model,
        "content": [{"type": "text", "text": content_text}],
        "stop_reason": stop_reason,
        "stop_sequence": None,
        "usage": {
            "input_tokens":  usage.get("prompt_tokens", 0),
            "output_tokens": usage.get("completion_tokens", 0),
        },
    }


def stream_openai_to_anthropic(openai_stream_resp, model: str):
    """
    将 DeepSeek SSE 流式响应转换为 Anthropic SSE 格式并 yield。

    Anthropic SSE 事件序列：
      message_start -> content_block_start -> content_block_delta(s)
      -> content_block_stop -> message_delta -> message_stop
    """
    # 发送 message_start
    yield f"event: message_start\ndata: {json.dumps({'type': 'message_start', 'message': {'id': 'msg_proxy', 'type': 'message', 'role': 'assistant', 'model': model, 'content': [], 'stop_reason': None, 'stop_sequence': None, 'usage': {'input_tokens': 0, 'output_tokens': 0}}})}\n\n"
    yield f"event: content_block_start\ndata: {json.dumps({'type': 'content_block_start', 'index': 0, 'content_block': {'type': 'text', 'text': ''}})}\n\n"
    yield f"event: ping\ndata: {json.dumps({'type': 'ping'})}\n\n"

    output_tokens = 0

    for line in openai_stream_resp.iter_lines():
        if not line:
            continue
        line = line.decode("utf-8") if isinstance(line, bytes) else line
        if not line.startswith("data: "):
            continue
        data_str = line[6:]
        if data_str.strip() == "[DONE]":
            break

        try:
            chunk = json.loads(data_str)
        except json.JSONDecodeError:
            continue

        choice = chunk.get("choices", [{}])[0]
        delta = choice.get("delta", {})
        text = delta.get("content", "")
        finish_reason = choice.get("finish_reason")

        if text:
            output_tokens += 1
            yield f"event: content_block_delta\ndata: {json.dumps({'type': 'content_block_delta', 'index': 0, 'delta': {'type': 'text_delta', 'text': text}})}\n\n"

        if finish_reason:
            stop_reason_map = {"stop": "end_turn", "length": "max_tokens"}
            stop_reason = stop_reason_map.get(finish_reason, "end_turn")
            yield f"event: content_block_stop\ndata: {json.dumps({'type': 'content_block_stop', 'index': 0})}\n\n"
            yield f"event: message_delta\ndata: {json.dumps({'type': 'message_delta', 'delta': {'stop_reason': stop_reason, 'stop_sequence': None}, 'usage': {'output_tokens': output_tokens}})}\n\n"
            yield f"event: message_stop\ndata: {json.dumps({'type': 'message_stop'})}\n\n"


@app.route("/v1/messages", methods=["POST"])
def messages():
    """
    主代理端点：接收 Anthropic /v1/messages 请求，转发到 DeepSeek。
    """
    if not DEEPSEEK_API_KEY:
        return jsonify({"error": "DEEPSEEK_API_KEY 未设置"}), 500

    body = request.get_json(force=True)
    is_stream = body.get("stream", False)

    # 模型映射
    claude_model = body.get("model", "claude-sonnet-4-5")
    ds_model = map_model(claude_model)
    logger.info(f"模型映射: {claude_model} -> {ds_model} | stream={is_stream}")

    # 构建 DeepSeek 请求体
    messages_list = anthropic_to_openai_messages(body)
    ds_body = {
        "model":    ds_model,
        "messages": messages_list,
        "max_tokens": min(int(body.get("max_tokens", 8192)), 8192),
        "stream":   is_stream,
    }

    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type":  "application/json",
    }

    # 修复：temperature 超出范围时 DeepSeek 返回 400，限制在 [0, 2]
    temp = float(body.get("temperature", 1.0))
    ds_body["temperature"] = max(0.0, min(2.0, temp))

    # 修复：top_p 如果存在也需要限制
    if "top_p" in body:
        ds_body["top_p"] = max(0.0, min(1.0, float(body["top_p"])))

    logger.debug(f"转发请求体: {json.dumps(ds_body, ensure_ascii=False)[:300]}")

    try:
        if is_stream:
            # 流式：透传 SSE，转换格式
            ds_resp = requests.post(
                f"{DEEPSEEK_API_BASE}/v1/chat/completions",
                headers=headers,
                json=ds_body,
                stream=True,
                timeout=120,
                proxies=REQUESTS_PROXIES,
            )
            if not ds_resp.ok:
                err = ds_resp.text
                logger.error(f"DeepSeek 流式错误 {ds_resp.status_code}: {err}")
                return jsonify({"error": {"type": "api_error", "message": err}}), 502
            return Response(
                stream_openai_to_anthropic(ds_resp, claude_model),
                content_type="text/event-stream",
                headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
            )
        else:
            # 非流式：转换响应格式
            ds_resp = requests.post(
                f"{DEEPSEEK_API_BASE}/v1/chat/completions",
                headers=headers,
                json=ds_body,
                timeout=120,
                proxies=REQUESTS_PROXIES,
            )
            if not ds_resp.ok:
                err = ds_resp.text
                logger.error(f"DeepSeek 错误 {ds_resp.status_code}: {err}")
                return jsonify({"error": {"type": "api_error", "message": err}}), 502
            anthropic_resp = openai_to_anthropic_response(ds_resp.json(), claude_model)
            return jsonify(anthropic_resp)

    except requests.exceptions.RequestException as e:
        logger.error(f"DeepSeek API 请求失败: {e}")
        return jsonify({"error": {"type": "api_error", "message": str(e)}}), 502


@app.route("/v1/models", methods=["GET"])
def models():
    """返回伪造的模型列表，让 Claude Code 认为模型可用"""
    return jsonify({
        "data": [
            {"id": "claude-sonnet-4-5",           "object": "model"},
            {"id": "claude-opus-4-5",             "object": "model"},
            {"id": "claude-3-5-sonnet-20241022",  "object": "model"},
            {"id": "claude-3-5-haiku-20241022",   "object": "model"},
        ]
    })


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "backend": "deepseek", "model": DEFAULT_DEEPSEEK_MODEL})


if __name__ == "__main__":
    if not DEEPSEEK_API_KEY:
        print("错误：请先设置 DEEPSEEK_API_KEY 环境变量")
        print("  Windows: $env:DEEPSEEK_API_KEY='sk-your-key'")
        sys.exit(1)

    print(f"DeepSeek 代理启动: http://{PROXY_HOST}:{PROXY_PORT}")
    print(f"默认模型: {DEFAULT_DEEPSEEK_MODEL}")
    print(f"请设置环境变量: ANTHROPIC_BASE_URL=http://{PROXY_HOST}:{PROXY_PORT}")
    print("然后启动 Claude Code 即可使用 DeepSeek 模型")

    app.run(host=PROXY_HOST, port=PROXY_PORT, debug=False)
