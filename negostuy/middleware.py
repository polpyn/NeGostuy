import sys
import time
from pathlib import Path

_LOG = Path(__file__).resolve().parent.parent / "requests.log"


def _w(line: str):
    try:
        with open(_LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
            f.flush()
    except Exception:
        pass
    try:
        print(line, flush=True)
    except Exception:
        pass


class RequestLogMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response
        _w(f"[middleware] сервер запущен, лог: {_LOG}")

    def __call__(self, request):
        t0 = time.time()
        _w(
            f"[→] {request.method} {request.get_full_path()}"
            f"  ip={request.META.get('REMOTE_ADDR','?')}"
            f"  host={request.META.get('HTTP_HOST','?')}"
        )
        response = self.get_response(request)
        ms = int((time.time() - t0) * 1000)
        _w(f"[←] {response.status_code}  {request.method} {request.get_full_path()}  {ms}ms")
        return response
