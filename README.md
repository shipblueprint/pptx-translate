# PPTX Translate

Translate PowerPoint (.pptx) files while preserving layout. Supports multiple providers (see PDF for quotas/limits).

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env
# Edit .env with your API keys/settings
```

## Usage

```bash
# Use settings from .env
python translate_presentation.py

# Or override
TRANSLATOR_PROVIDER=gtx TARGET_LANG=ZH INPUT_PPTX=Pitch-Deck.pptx python translate_presentation.py
```

## Providers

- `deepl` - DeepL API (needs `DEEPL_API_KEY`)
- `deeplx` - Self-hosted/free proxy (e.g., DeepLX)
- `gtx` - Google Translate unofficial endpoint (no key usually)
- `libretranslate` - LibreTranslate/self-hosted
- `qwen_dashscope` / `openai_compatible` - Qwen/OpenRouter/etc via OpenAI-compatible API

See `add the charachetr limit qouta to the cloud optio....pdf` for full comparison.

## Fallback chain

By default it tries providers in order until one succeeds. Configure via:

```bash
TRANSLATOR_CHAIN=deeplx,gtx,libretranslate,qwen_dashscope
```

Or use a single provider:
```bash
TRANSLATOR_PROVIDER=deepl
```

It logs which provider handled each chunk and warns when approaching quota. If all fail, it keeps original text.
