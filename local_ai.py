from __future__ import annotations

import os
from dataclasses import dataclass

import requests


@dataclass
class LocalAIResult:
    ok: bool
    message: str


def ask_local_ai(instruction: str, file_name: str, file_text: str) -> LocalAIResult:
    base_url = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434").rstrip("/")
    model = os.getenv("OLLAMA_MODEL", "llama3.1:8b")

    if not file_text.strip():
        return LocalAIResult(False, "Fayldan mətn çıxara bilmədim. Əgər PDF skandırsa, OCR aktiv olmalıdır.")

    prompt = (
        "Sən Azərbaycan dilində danışan peşəkar ofis assistentisən. "
        "İstifadəçi sənə bir ofis faylı və tapşırıq verir. "
        "Fayldakı məlumata əsaslan, uydurma məlumat yazma. "
        "Əgər tapşırıq faylı redaktə etməyi tələb edirsə, hansı dəyişiklikləri etməli olduğunu dəqiq siyahıla. "
        "Cavabı qısa, praktik və Azərbaycan dilində ver.\n\n"
        f"Fayl adı: {file_name}\n"
        f"Tapşırıq: {instruction}\n\n"
        "Fayldan çıxarılan məzmun:\n"
        f"{_trim_context(file_text)}"
    )

    try:
        response = requests.post(
            f"{base_url}/api/chat",
            json={
                "model": model,
                "messages": [{"role": "user", "content": prompt}],
                "stream": False,
            },
            timeout=120,
        )
        response.raise_for_status()
    except requests.RequestException as exc:
        return LocalAIResult(
            False,
            "Lokal AI işləmədi. Ollama quraşdırılıb başladılmalıdır. "
            f"Texniki xəta: {exc}",
        )

    data = response.json()
    content = data.get("message", {}).get("content", "").strip()
    if not content:
        return LocalAIResult(False, "Lokal AI boş cavab qaytardı.")

    return LocalAIResult(True, content)


def _trim_context(text: str, limit: int = 18000) -> str:
    compact = text.strip()
    if len(compact) <= limit:
        return compact
    return compact[:limit].rstrip() + "\n\n[Mətn çox uzun olduğu üçün davamı kəsildi.]"

