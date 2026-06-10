# Codex Repository Instructions

Read `docs/CODEX_HANDOVER.md` before making changes. It contains the current product status, recent work, and known issues.

## Working Style

- The user wants changes implemented proactively, not left as proposals.
- After completing each requested change:
  1. Run `npm.cmd run build`.
  2. Run `node --check server/index.js` when the server was changed.
  3. Stage only the files changed for the request.
  4. Commit with a concise message.
  5. Push the current branch automatically.
- Do not ask the user to run PowerShell for ordinary builds, commits, or pushes.
- Keep the user updated while working and report the pushed commit hash when finished.

## Git And Data Safety

- The normal working branch is `main`, pushed to `origin`.
- Never use `git add .` or stage unrelated files.
- Never overwrite, revert, or commit local operational data unless the user explicitly requests it.
- In particular, leave existing changes in `data/jobs.json` and `data/users.json` alone.
- Ignore OneDrive conflict copies such as files containing `-LAPTOP-`.
- Ignore local backups, generated inspection folders, and temporary PDF folders.
- Never use destructive Git commands such as `git reset --hard` or `git checkout --`.

## Main Application Files

- `src/App.jsx`: React UI and module behaviour.
- `src/index.css`: portal styling.
- `server/index.js`: Express API, CoreBridge integration, and persisted board data.
- `data/jobs.json`: local/live-style operational data; treat carefully.
- `data/users.json`: local user data; treat carefully.

## Product Expectations

- Match the existing portal visual language and keep operational screens clean, compact, and modern.
- The user prefers automatic deployment through commits pushed to `main`.
- Design Board and Filtering Board cards should remain visually consistent.
- Do not add hard-coded overrides for CoreBridge line item names or categories. Pull the exact source fields.

