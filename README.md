# Screen Translator

Press **Ctrl+Shift+T**, drag a box over any text on your screen, and get:

- **Original text** — the OCR result, layout and paragraph breaks preserved, fully editable
- **Translation** — into any of 14 languages, re-translates your edits
- **Summary** — 2–3 sentences in the target language (local AI via Ollama, or a built-in key-sentence summarizer)
- **Table for Excel** — tables are detected and copied as tab-separated cells that paste straight into Excel
- **Image** — the captured region, copy to clipboard or save as PNG

Works on anything visible: PDFs, websites, apps, images. **Free and fully offline** after one-time model downloads.

## Languages

English, Hindi, Spanish, French, German, Portuguese, Italian, Russian, Arabic, Japanese, Korean, Chinese (Simplified), Turkish, Urdu.
Translation between two non-English languages pivots through English automatically.

## Install

Requires Python 3.10+ from [python.org](https://www.python.org/) (Tkinter included). On Windows:

```bash
py -3.12 -m pip install -r requirements.txt
```

Linux: `python3 -m pip install -r requirements.txt` (plus `sudo apt install python3-tk`; for a smaller download install the CPU build of PyTorch first — see `requirements.txt`).

## Run

**Option A — always ready (recommended).** Installs a tiny hotkey listener (~35 MB RAM) that starts at every login. Ctrl+Shift+T then launches the app straight into capture mode:

```bash
py -3.12 screen_ocr_translator.py --install
```

Undo anytime with `--uninstall`. If you move the folder, run `--install` again.

**Option B — run it directly:**

```bash
py -3.12 screen_ocr_translator.py
```

### Using it

1. Press **Ctrl+Shift+T** — the screen dims.
2. Drag over the text. **Esc** or right-click cancels.
3. The results window opens. Change **Text in image** to re-run OCR, **Translate to** to re-translate, or edit the original and click **Translate again**.
4. Drag the gaps between boxes to resize them. Closing the window hides it (next capture is instant); **Quit** stops the app.

## First run

| Download | Size | When |
|---|---|---|
| EasyOCR models | ~100 MB + ~20 MB per extra language | first capture in a language |
| Argos Translate packs | ~100 MB per language pair | first translation of a pair |
| Ollama + `llama3.2` (optional) | ~2 GB | only if you want AI summaries: install [Ollama](https://ollama.com), run `ollama pull llama3.2` |

Set `OLLAMA_MODEL = None` at the top of the script to always use the built-in summarizer.

## Requirements

- ~1.5 GB disk without the AI summary, ~4 GB with it
- ~1–1.5 GB RAM while the app is open (no GPU needed; an NVIDIA GPU is used automatically if present)
- Works on low-end laptops — just slower (first capture ~10–20 s, then a few seconds)

## Troubleshooting

- **Selection box doesn't line up** — DPI scaling issue; use python.org Python (not MSYS/Store builds).
- **Black capture** — fullscreen games and DRM video players can't be captured; use windowed mode.
- **Slow first capture** — the OCR model is loading/downloading; later captures are much faster.
- **Hotkey does nothing** — when an app running as administrator is focused, Windows blocks the hotkey.
- **Logs** — `%TEMP%\screen_ocr_translator.log`
- **Linux** — the global hotkey needs X11 (not Wayland); Copy image needs `xclip` or `wl-clipboard`.

## Self-test

Checks the layout, table, summarizer and translation logic without GUI or downloads:

```bash
py -3.12 screen_ocr_translator.py --selftest
```

## Built with

[EasyOCR](https://github.com/JaidedAI/EasyOCR) · [Argos Translate](https://github.com/argosopentech/argos-translate) · [mss](https://github.com/BoboTiG/python-mss) · [Pillow](https://python-pillow.org/) · [pynput](https://github.com/moses-palmer/pynput) · Tkinter

## License

[MIT](LICENSE) © 2026 SARTHAK8377
