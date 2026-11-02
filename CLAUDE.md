# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Communication

- Always reply to the user in Russian. Leave code, code comments, UI strings and commit messages in their existing language (Ukrainian); don't translate them.

## Source vs generated files

- Edit `template_v2.html` (the single-file vanilla HTML/CSS/JS app; Ukrainian UI). Never hand-edit `index.html`, because it is generated.
- Build: `python generate_html.py` injects `tabel_data.json` into the `%%EMPLOYEES%%` / `%%SECTIONS%%` / `%%OBJECTS%%` placeholders and writes `index.html`, which is served by GitHub Pages and committed. Re-run it after every template or data change.
- `Табель_2026.html` is the legacy output name. Don't update it.
- The `v*/`, `финал*/` and `github_upload/` folders are frozen release snapshots. Don't edit them, and only create a new one when the user asks. The format is `vYYMMDDHHMM/` with copies of `index.html`, `template_v2.html`, `apps_script.gs`, `generate_html.py`, `tabel_data.json` plus a `CHANGELOG.md`.

## Backend (Google Apps Script)

- `apps_script.gs` is the backend (`doGet`/`doPost`, writes Sheets columns A–S, sends Telegram). It contains the hardcoded Telegram token and is gitignored. Never commit it or paste its secrets anywhere, including its copies inside snapshot folders.
- After changing `apps_script.gs`, remind the user to paste it into the Apps Script editor, then go to Deploy → Manage deployments → ✏️ edit the existing deployment → Version: **New version** → Deploy. This keeps the `/exec` URL unchanged, so don't touch `DEFAULT_URL`. Never suggest "New deployment": it creates a new URL while the old one keeps running old code for cached pages.
- The frontend POSTs with `mode:'no-cors'` + `Content-Type: text/plain`, so the client cannot read the response.

## Versioning

- On every release, bump the version (format `vYY.MM.DD`) in all four places: `#appVersion` in `template_v2.html`, and in `apps_script.gs` the header comment, the `doGet` "Табель API vXX працює" message and the `testTelegram` message.
- Release commits are in Ukrainian and use the form `vYYMMDDHHMM — <summary>`, followed by a bullet list.

## Gotchas

- `generate_html.py` writes in text mode, so on Windows `index.html` gets CRLF while the template is LF. Whole-file diffs in `index.html` are expected.
- The project path contains Cyrillic, so always quote paths in shell commands.
