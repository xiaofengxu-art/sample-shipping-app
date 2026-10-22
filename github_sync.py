"""在线上（Streamlit Cloud）把模板文件提交回 GitHub 仓库。

云端容器的文件系统是临时的，上传的文件重启就没了，
所以线上更新模板必须写回仓库：提交后 Streamlit Cloud 会自动重新部署。
"""

from __future__ import annotations

import base64
import json
import urllib.error
import urllib.parse
import urllib.request

API = "https://api.github.com"


def _call(url: str, token: str, method: str = "GET", payload: dict | None = None) -> dict:
    data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("User-Agent", "sample-shipping-app")
    if data is not None:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode() or "{}")
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return {"__missing__": True}
        detail = exc.read().decode(errors="ignore")[:200]
        raise RuntimeError(f"GitHub 返回 {exc.code}：{detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"连不上 GitHub：{exc.reason}") from exc


def replace_template(
    repo: str,
    path: str,
    content: bytes,
    token: str,
    branch: str = "main",
    message: str = "页面上传更新商品模板",
) -> str:
    """用 content 覆盖仓库中的 path，返回本次提交的 sha。"""
    encoded = urllib.parse.quote(path)
    meta = _call(f"{API}/repos/{repo}/contents/{encoded}?ref={branch}", token)
    payload: dict[str, str] = {
        "message": message,
        "content": base64.b64encode(content).decode(),
        "branch": branch,
    }
    if not meta.get("__missing__"):
        payload["sha"] = meta["sha"]
    result = _call(f"{API}/repos/{repo}/contents/{encoded}", token, method="PUT", payload=payload)
    return result.get("commit", {}).get("sha", "")
