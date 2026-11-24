"""Shared helpers: Google Gemini (free tier) and Jina Reader (free page reader)."""
from __future__ import annotations

import json
import os
import re
import time

import requests

from .common import load_config, load_env

GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent"


def gemini_key() -> str:
    load_env()
    return os.environ.get("GEMINI_API_KEY", "").strip()


def gemini(prompt: str, json_mode: bool = False, max_tokens: int = 4000) -> tuple[str, str]:
    """Return (text, model). Tries each configured free model in turn."""
    cfg = load_config().get("gemini", {})
    key = gemini_key()
    if not key:
        raise RuntimeError("GEMINI_API_KEY not set")
    gen = {"temperature": 0.1 if json_mode else 0.3, "maxOutputTokens": max_tokens}
    if json_mode:
        gen["responseMimeType"] = "application/json"
    last = ""
    for m in cfg.get("models", ["gemini-flash-latest"]):
        for attempt in range(2):
            r = requests.post(GEMINI_URL.format(m=m),
                              headers={"x-goog-api-key": key, "Content-Type": "application/json"},
                              json={"contents": [{"parts": [{"text": prompt}]}], "generationConfig": gen},
                              timeout=90)
            if r.status_code == 200:
                try:
                    txt = r.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
                    time.sleep(cfg.get("pause_seconds", 5))
                    return txt, m
                except (KeyError, IndexError):
                    last = f"{m}: empty answer"
                    break
            last = f"{m}: HTTP {r.status_code} {r.text[:160]}"
            if r.status_code == 429 and attempt == 0:
                time.sleep(30)
                continue
            break
    raise RuntimeError(last)


def gemini_json(prompt: str):
    txt, model = gemini(prompt, json_mode=True)
    txt = re.sub(r"^```(?:json)?|```$", "", txt.strip(), flags=re.M).strip()
    return json.loads(txt), model


_cache: dict[str, str] = {}
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36"}


def resolve(url: str) -> str:
    """Google AI Mode cites sources as google.com/goto?url=... redirects; return the real page URL."""
    if not url or "google." not in url.split("/")[2:3][0] if url.startswith("http") else True:
        return url
    try:
        r = requests.get(url, allow_redirects=False, timeout=20, headers=UA)
        loc = r.headers.get("location", "")
        return loc if loc.startswith("http") else url
    except requests.RequestException:
        return url


def _direct(url: str) -> str:
    try:
        r = requests.get(url, timeout=20, headers=UA)
        if r.status_code != 200 or "html" not in r.headers.get("content-type", "html"):
            return ""
        html = re.sub(r"(?is)<(script|style|noscript)[^>]*>.*?</\1>", " ", r.text)
        title = re.search(r"(?is)<title[^>]*>(.*?)</title>", html)
        text = re.sub(r"<[^>]+>", " ", html)
        text = re.sub(r"&nbsp;|&#160;", " ", text)
        return (f"Title: {title.group(1).strip()}\n" if title else "") + re.sub(r"\s+", " ", text)
    except requests.RequestException:
        return ""


def read_page(url: str, timeout: int = 25) -> str:
    """Page text via Jina Reader (handles JavaScript sites and many bot blocks). '' on failure."""
    if not url:
        return ""
    if url in _cache:
        return _cache[url]
    text = _direct(url)
    if len(text) > 1500:
        _cache[url] = text
        return text
    headers = {"Accept": "text/plain", "X-Return-Format": "text"}
    if os.environ.get("JINA_API_KEY"):
        headers["Authorization"] = f"Bearer {os.environ['JINA_API_KEY']}"
    try:
        r = requests.get("https://r.jina.ai/" + url, headers=headers, timeout=timeout)
        if r.status_code == 200 and len(r.text) > len(text):
            text = r.text
    except requests.RequestException:
        pass
    time.sleep(3)  # free Jina tier is rate limited
    _cache[url] = text
    return text


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^\w€£.,%-]+", " ", (s or "").lower())).strip()


def page_supports(page: str, value, quote: str = "") -> bool:
    """True if the page text backs the value: numbers must appear (allowing 1,225 / 1.225 / 1225),
    text values or the supporting quote must mostly appear."""
    if not page:
        return False
    p = norm(page)
    v = str(value).strip()
    num = re.fullmatch(r"\d+(?:\.\d+)?", v.replace(",", ""))
    if num:
        n = float(v.replace(",", ""))
        forms = {f"{n:g}", f"{n:,.0f}", f"{n:,.0f}".replace(",", "."), f"{n:,.0f}".replace(",", " ")}
        if n >= 1000 and n % 1000 == 0:
            forms |= {f"{n/1000:g} billion", f"{n/1000:g}bn", f"{n/1000:g} bn", f"{n/1000:g} gw"}
        if n >= 1000:
            forms |= {f"{n/1000:.1f} billion", f"{n/1000:.1f}bn", f"{n/1000:.2f} billion".rstrip("0")}
        return any(f.lower() in p for f in forms)
    if re.fullmatch(r"\d{4}(-\d{2}(-\d{2})?)?", v):
        return v[:4] in p
    words = [w for w in norm(quote or v).split() if len(w) > 3]
    if not words:
        return False
    hit = sum(1 for w in words if w in p)
    return hit / len(words) >= 0.7


def page_title(page: str, url: str) -> str:
    m = re.search(r"^Title:\s*(.+)$", page or "", re.M)
    return (m.group(1).strip() if m else url)[:140]
