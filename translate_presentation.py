#!/usr/bin/env python3
"""Translate PPTX files using multiple providers with fallback chain while preserving layout."""

from __future__ import annotations

import json
import os
import shutil
import sys
from typing import Callable, Optional

import httpx
from dotenv import load_dotenv
from pptx import Presentation

load_dotenv()

# Config
PROVIDER = os.getenv("TRANSLATOR_PROVIDER", "").lower().strip()
CHAIN = os.getenv("TRANSLATOR_CHAIN", "").strip()
TARGET_LANG = os.getenv("TARGET_LANG", "ZH")
INPUT_PPTX = os.getenv("INPUT_PPTX", "Pitch-Deck.pptx")
OUTPUT_PPTX = os.getenv("OUTPUT_PPTX", f"{os.path.splitext(INPUT_PPTX)[0]}_translated.pptx")
QUOTA_WARN = int(os.getenv("QUOTA_WARN_CHARS", "400000"))

# Stats
_char_count = 0
_provider_usage = {}
_skipped_config = set()


def get_chain() -> list[str]:
    if CHAIN:
        return [p.strip().lower() for p in CHAIN.split(",") if p.strip()]
    if PROVIDER:
        return [PROVIDER]
    # sensible defaults
    return ["deeplx", "gtx", "libretranslate", "deepl"]


def normalize_text(text: str) -> str:
    return text.replace("\x0b", "").replace("\r", "").strip()


def iter_text_frames(shapes):
    for shape in shapes:
        if shape.shape_type == 6:  # GROUP
            yield from iter_text_frames(shape.shapes)
        elif hasattr(shape, "text_frame") and shape.has_text_frame:
            yield shape, shape.text_frame


def replace_in_paragraph(p, new_text: str):
    runs = p.runs
    if not runs:
        p.add_run().text = new_text
        return
    runs[0].text = new_text
    for r in runs[1:]:
        try:
            r._r.getparent().remove(r._r)
        except Exception:
            r.text = ""


# Providers
_deepl_translator = None


def translate_deepl(text: str) -> str:
    global _deepl_translator
    key = os.environ.get("DEEPL_API_KEY")
    if not key:
        raise RuntimeError("DEEPL_API_KEY not set")
    import deepl

    if _deepl_translator is None:
        _deepl_translator = deepl.Translator(key)
    result = _deepl_translator.translate_text(text, target_lang=TARGET_LANG)
    return result.text


def translate_deeplx(text: str) -> str:
    url = os.getenv("DEEPLX_URL", "http://localhost:1188/translate")
    resp = httpx.post(url, json={"text": text, "source_lang": "auto", "target_lang": TARGET_LANG}, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    return data.get("data") or data.get("text") or text


def translate_gtx(text: str) -> str:
    url = "https://translate.googleapis.com/translate_a/single"
    resp = httpx.get(url, params={"client": "gtx", "sl": "auto", "tl": TARGET_LANG, "dt": "t", "q": text}, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    try:
        return "".join(part[0] for part in data[0] if part and part[0])
    except Exception:
        return text


def translate_edge(text: str) -> str:
    url = os.getenv("EDGE_URL", "https://api-edge.cognitive.microsofttranslator.com/translate")
    resp = httpx.post(
        url,
        params={"api-version": "3.0", "from": "auto", "to": TARGET_LANG.lower()},
        headers={"Content-Type": "application/json"},
        json=[{"Text": text}],
        timeout=60,
    )
    resp.raise_for_status()
    data = resp.json()
    try:
        return data[0]["translations"][0]["text"]
    except Exception:
        return text


def translate_google(text: str) -> str:
    key = os.environ.get("GOOGLE_API_KEY")
    if not key:
        raise RuntimeError("GOOGLE_API_KEY not set")
    resp = httpx.post(
        f"https://translation.googleapis.com/language/translate/v2?key={key}",
        json={"q": [text], "target": TARGET_LANG.lower(), "format": "text"},
        timeout=60,
    )
    resp.raise_for_status()
    data = resp.json()
    try:
        return data["data"]["translations"][0]["translatedText"]
    except Exception:
        return text


def translate_googlecloud(text: str) -> str:
    try:
        from google.cloud import translate_v2 as translate

        client = translate.Client()
        result = client.translate(text, target_language=TARGET_LANG.lower())
        return result["translatedText"]
    except Exception as e:
        raise RuntimeError(f"google cloud translate failed: {e}")


def translate_azure(text: str) -> str:
    key = os.environ.get("AZURE_TRANSLATOR_KEY")
    region = os.environ.get("AZURE_TRANSLATOR_REGION")
    if not key or not region:
        raise RuntimeError("AZURE_TRANSLATOR_KEY and AZURE_TRANSLATOR_REGION not set")
    endpoint = os.getenv("AZURE_TRANSLATOR_ENDPOINT", "https://api.cognitive.microsofttranslator.com")
    resp = httpx.post(
        f"{endpoint}/translate",
        params={"api-version": "3.0", "to": TARGET_LANG.lower()},
        headers={"Ocp-Apim-Subscription-Key": key, "Ocp-Apim-Subscription-Region": region, "Content-Type": "application/json"},
        json=[{"Text": text}],
        timeout=60,
    )
    resp.raise_for_status()
    data = resp.json()
    try:
        return data[0]["translations"][0]["text"]
    except Exception:
        return text


def translate_amazon(text: str) -> str:
    try:
        import boto3

        region = os.getenv("AWS_REGION", "us-east-1")
        client = boto3.client("translate", region_name=region)
        resp = client.translate_text(Text=text, SourceLanguageCode="auto", TargetLanguageCode=TARGET_LANG.lower())
        return resp["TranslatedText"]
    except Exception as e:
        raise RuntimeError(f"amazon translate failed: {e}")


def translate_libretranslate(text: str) -> str:
    url = os.getenv("LIBRETRANSLATE_URL", "http://localhost:5000/translate")
    payload = {"q": text, "source": "auto", "target": TARGET_LANG.lower(), "format": "text"}
    api_key = os.getenv("LIBRETRANSLATE_API_KEY")
    if api_key:
        payload["api_key"] = api_key
    resp = httpx.post(url, json=payload, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    return data.get("translatedText", text)


def translate_openai_compatible(text: str) -> str:
    base_url = os.getenv("OPENAI_BASE_URL") or os.getenv("QWEN_BASE_URL") or os.getenv("DASHSCOPE_BASE_URL")
    api_key = os.getenv("OPENAI_API_KEY") or os.getenv("DASHSCOPE_API_KEY")
    model = os.getenv("TRANSLATE_MODEL", "qwen-mt-turbo")
    if not base_url:
        raise RuntimeError("OPENAI_BASE_URL/DASHSCOPE_BASE_URL required")
    resp = httpx.post(
        f"{base_url.rstrip('/')}/chat/completions",
        headers={"Authorization": f"Bearer {api_key}" if api_key else "", "Content-Type": "application/json"},
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": f"Translate the following text to {TARGET_LANG}. Return only the translated text with no explanation."},
                {"role": "user", "content": text},
            ],
            "temperature": 0,
        },
        timeout=180,
    )
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"].strip()


TRANSLATORS: dict[str, Callable[[str], str]] = {
    "deepl": translate_deepl,
    "deeplx": translate_deeplx,
    "gtx": translate_gtx,
    "edge": translate_edge,
    "google": translate_google,
    "googlecloud": translate_googlecloud,
    "azure": translate_azure,
    "amazon": translate_amazon,
    "libretranslate": translate_libretranslate,
    "openai_compatible": translate_openai_compatible,
    "qwen_dashscope": translate_openai_compatible,
    "qwen": translate_openai_compatible,
    "translategemma": translate_openai_compatible,
    "ollama": translate_openai_compatible,
}


def translate_text(text: str) -> str:
    global _char_count
    cleaned = normalize_text(text)
    if not cleaned:
        return text
    _char_count += len(cleaned)
    if _char_count >= QUOTA_WARN and _char_count - len(cleaned) < QUOTA_WARN:
        print(f"⚠️  Approaching/quota warning: { _char_count } chars processed (warn at {QUOTA_WARN})", file=sys.stderr)

    for name in get_chain():
        fn = TRANSLATORS.get(name)
        if not fn:
            print(f"  [skip] unknown provider '{name}'", file=sys.stderr)
            continue
        try:
            res = fn(cleaned)
            if res and res.strip() and res.strip() != cleaned:
                _provider_usage[name] = _provider_usage.get(name, 0) + len(cleaned)
                return res
            # if same as original, try next
        except Exception as e:
            print(f"  [{name}] failed: {e}", file=sys.stderr)
            continue
    print(f"  [all failed] keeping original: '{cleaned[:40]}...'", file=sys.stderr)
    return text


def apply_replacements(tf):
    for p in tf.paragraphs:
        original = normalize_text(p.text)
        if original:
            translated = translate_text(original)
            replace_in_paragraph(p, translated)


def translate_presentation():
    if not os.path.exists(INPUT_PPTX):
        print(f"Error: {INPUT_PPTX} not found", file=sys.stderr)
        sys.exit(1)

    chain = get_chain()
    print(f"Provider chain: {chain} | Target: {TARGET_LANG} | Input: {INPUT_PPTX}")
    backup_path = f"{os.path.splitext(INPUT_PPTX)[0]}_backup.pptx"
    shutil.copy(INPUT_PPTX, backup_path)
    print(f"Created backup: {backup_path}")

    prs = Presentation(INPUT_PPTX)
    print(f"Translating {len(prs.slides)} slides...")

    for slide_idx, slide in enumerate(prs.slides, 1):
        print(f"  Processing Slide {slide_idx}...")
        for shape, tf in iter_text_frames(slide.shapes):
            apply_replacements(tf)

    prs.save(OUTPUT_PPTX)
    print(f"\n✓ Translation complete!")
    print(f"  Original: {INPUT_PPTX}")
    print(f"  Backup: {backup_path}")
    print(f"  Translated: {OUTPUT_PPTX}")
    print(f"  Total chars processed: {_char_count}")
    if _provider_usage:
        print("  Provider usage (chars):")
        for k, v in _provider_usage.items():
            print(f"    - {k}: {v}")


if __name__ == "__main__":
    translate_presentation()
