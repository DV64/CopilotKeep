# CopilotKeep — M365 chat backup & export

[![Version](https://img.shields.io/badge/version-0.1.0-6366f1)](#changelog)
[![License: MIT](https://img.shields.io/badge/license-MIT-10b981)](./LICENSE)

A Tampermonkey userscript that backs up and exports your Microsoft 365 Copilot chats — the current conversation or your whole archive — to Markdown, JSON, ChatGPT-compatible JSON, or a ZIP bundle.

> **Unofficial tool.** Not affiliated with or endorsed by Microsoft. It uses undocumented internal endpoints and may stop working if Microsoft changes them.

## Features

- **Current chat** — copy or download the visible conversation (MD / HTML / JSON), or fetch the complete history fresh from Microsoft's API
- **Archive** — list, search, sort, and batch-export conversations (deflate-compressed ZIP or combined files)
- **Formats** — Markdown, Clean JSON, ChatGPT-compatible JSON, raw API JSON (behind an explicit confirmation)
- **Settings tab** — interface language (**English / العربية / 中文**, with RTL support), theme (auto / light / dark), export defaults, fetch pacing, page size
- **Robust fetching** — `Retry-After` backoff on 429/503, automatic token refresh on 401, partial-view detection for virtualized long chats
- **Read-only by design** — no delete operations

## Installation

1. Install [Tampermonkey](https://www.tampermonkey.net/) (Chrome, Edge, Firefox, …).
2. Install the script, either:
   - open the raw file URL — **[CopilotKeep — M365 chat backup & export](https://raw.githubusercontent.com/DV64/CopilotKeep/main/copilotkeep.user.js)** — and Tampermonkey offers a one-click install, or
   - create a new script in the Tampermonkey dashboard and paste the file contents.
3. Open `https://m365.cloud.microsoft/` or `https://copilot.cloud.microsoft/` — a **↓** launcher button appears at the bottom-right corner.

## Usage

- **Current chat tab** — shows the conversation state (visible messages, whether full history is cached, partial-view warnings). *Page* buttons read only the visible DOM; **Export from Microsoft** always fetches fresh, complete history.
- **Archive tab** — click **Load** to list your recent conversations, search/sort, select, choose a format, then **Export selected**. Progress and per-row status are shown inline.
- **Logs tab** — filterable operation log (info / warn / error) with copy and clear.
- **Settings tab** — language, theme, default formats, ZIP compression toggle, request delay, page size, and reset to defaults.

## Privacy & security

- Everything runs locally in your browser. Requests go **only** to Microsoft's own endpoints (`substrate.office.com`) with your existing session — no third-party calls, no analytics, no telemetry.
- To authenticate archive/full exports, the script obtains a Bearer token by: (1) capturing a token from the page's own network traffic, or (2) as a fallback, decrypting the MSAL token cache in `localStorage` using the page's `msal.cache.encryption` cookie. Tokens stay in memory and are never sent anywhere except Microsoft.
- Exports may contain personal data — handle the downloaded files accordingly.
- **Raw JSON** exports can include IDs, tenant information, and telemetry fields; the script warns before exporting them.

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| "Encryption cookie missing" / "No valid substrate token" | Reload the Copilot page (session tokens rotate), then retry. |
| Exports stall periodically | The script honors `Retry-After` on rate limits; raise **Delay between fetches** in Settings. |
| Long chat exports are incomplete | *Page* methods see only the virtualized window — use **Export from Microsoft** for full history. |
| Archive stops loading after a Microsoft update | The internal API changed; check for an updated script version. |

## Changelog

See [CHANGELOG.md](./CHANGELOG.md). Current version: **0.1.0** (2026-11-12) — first release.

## License

[MIT](./LICENSE) © DV64
