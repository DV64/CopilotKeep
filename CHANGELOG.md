# Changelog

All notable changes to **CopilotKeep** are documented here.
Format based on [Keep a Changelog](https://keepachangelog.com/); versioning follows [SemVer](https://semver.org/).

## [0.1.0] — 2026-11-12

First release.

- Export the current Copilot chat: copy to clipboard, or download as Markdown / HTML / JSON
- Full export that fetches the complete history fresh from Microsoft's API
- Archive tab: list, search, sort, and batch-export conversations
- Formats: Markdown, Clean JSON, ChatGPT-compatible JSON, raw API JSON
- Batch export to a deflate-compressed ZIP (STORE fallback)
- Settings tab: interface language (English / العربية / 中文, with RTL), theme (auto / light / dark), export defaults, fetch pacing, page size
- Rate-limit handling (`Retry-After` backoff on 429/503) and automatic token refresh on 401
- Partial-view detection for long virtualized chats
- Read-only by design: no delete operations