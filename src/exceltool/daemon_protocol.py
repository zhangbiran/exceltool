import json
import struct

from .errors import ExcelToolError


PROTOCOL_VERSION = 1
MAX_FRAME_SIZE = 64 * 1024
HEADER = struct.Struct("!I")


class ProtocolError(ExcelToolError):
    def __init__(self, message, error_code="bad_request"):
        super().__init__(message, 5)
        self.error_code = error_code


def _receive_exact(sock, size):
    chunks = []
    remaining = size
    while remaining:
        chunk = sock.recv(remaining)
        if not chunk:
            if remaining == size:
                return None
            raise ProtocolError("daemon 控制连接在消息中途关闭")
        chunks.append(chunk)
        remaining -= len(chunk)
    return b"".join(chunks)


def receive_message(sock):
    header = _receive_exact(sock, HEADER.size)
    if header is None:
        return None
    size = HEADER.unpack(header)[0]
    if size == 0 or size > MAX_FRAME_SIZE:
        raise ProtocolError("daemon 控制消息长度无效")
    payload = _receive_exact(sock, size)
    if payload is None:
        raise ProtocolError("daemon 控制连接缺少消息正文")
    try:
        message = json.loads(payload.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ProtocolError("daemon 控制消息不是有效 UTF-8 JSON: %s" % exc)
    if not isinstance(message, dict):
        raise ProtocolError("daemon 控制消息必须是 JSON 对象")
    return message


def send_message(sock, message):
    try:
        payload = json.dumps(
            message, ensure_ascii=False, separators=(",", ":")
        ).encode("utf-8")
    except (TypeError, ValueError) as exc:
        raise ProtocolError("无法编码 daemon 控制消息: %s" % exc)
    if not payload or len(payload) > MAX_FRAME_SIZE:
        raise ProtocolError("daemon 控制消息超过 %d 字节" % MAX_FRAME_SIZE)
    sock.sendall(HEADER.pack(len(payload)) + payload)


def require_common_request(message):
    message_type = message.get("type")
    if not isinstance(message_type, str) or not message_type:
        raise ProtocolError("daemon 控制消息缺少 type")
    if message.get("protocol_version") != PROTOCOL_VERSION:
        raise ProtocolError("daemon 控制协议版本不兼容", "incompatible_protocol")
    instance_id = message.get("instance_id")
    token = message.get("token")
    if not isinstance(instance_id, str) or not instance_id:
        raise ProtocolError("daemon 控制消息缺少 instance_id")
    if not isinstance(token, str) or not token:
        raise ProtocolError("daemon 控制消息缺少 token")
    return message_type, instance_id, token


def error_message(code, message):
    return {"type": "error", "code": code, "message": message}
