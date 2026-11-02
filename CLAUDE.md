# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Communication

- Always reply to the user in Russian. Leave code, code comments, UI strings and commit messages in their existing language (Ukrainian); don't translate them.

## Source vs generated files

- Edit `template_v2.html` (the single-file vanilla HTML/CSS/JS app; Ukrainian UI). Never hand-edit `index.html`, because it is generated.
- `tabel_data.json` is an offline fallback snapshot of the Sheets reference lists (employees, sections, objects, worktypes). The live form prefers `GET <DEFAULT_URL>?action=getData`, so refresh this file from that endpoint when the lists drift.
- Build: `python generate_html.py` injects `tabel_data.json` into the `%%EMPLOYEES%%` / `%%SECTIONS%%` / `%%OBJECTS%%` placeholders and writes `index.html`, which is served by GitHub Pages and committed. Re-run it after every template or data change.
- `generate_html.py` also stamps `%%BUILD%%` (YYMMDDHHMM) into `APP_BUILD`. Open pages poll the live `index.html` and show the "Вийшла нова версія" bar when the deployed build is newer, so always regenerate before publishing.
- Old release snapshots live in `_archive/` (gitignored, local only). Some hold the only copies of old `apps_script.gs` versions, so don't edit or delete them. Git history plus the GitHub remote (`origin`, branch `main`; GitHub Pages serves `index.html`) is now the release record, so don't create new snapshot folders unless the user asks.
- Publishing a form change means commit + `git push`, and Pages goes live in about 1 minute. Never force-push.

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
