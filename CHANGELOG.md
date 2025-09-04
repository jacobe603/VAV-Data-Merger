# Changelog

All notable changes to this project are documented in this file.

## 2025-09-04

### Added
- Original path override in `POST /refresh_and_compare` via `original_path` request field; also saved to session when valid.
- Structured logging (`logging` module) with INFO-level output; replaces ad-hoc prints in refresh/compare routes.
- mtime-based skip for refresh to avoid re-reading unchanged TW2 files; response includes `skipped_read` and `path_source`.

### Changed
- Standardized API responses for `/compare_performance` and `/refresh_and_compare` to a consistent shape:
  - Success: `{ success: true, data: { ... } }`
  - Error: `{ success: false, error: '...' }`
- Extracted large inline script from `templates/index.html` to `static/app.js` and updated the template to include it.
- Updated frontend fetch handlers to read from `data.data` and to pass `original_path` during refresh.

### Notes
- If any external tools consumed these endpoints directly, ensure they read from `data` on success and `error` on failure.
- Refresh endpoint now prefers the provided original file path when accessible; otherwise falls back to the local uploaded copy.

