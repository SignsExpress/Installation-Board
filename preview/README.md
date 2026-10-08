# Isolated portal draft

Rebuild work stays on `codex/portal-rebuild`. Do not merge or deploy to production until the user authorizes the switch.

The pro-forma PDF output is an explicit preservation requirement. Keep the existing `buildProFormaPreviewHtml`, calculations, print styles, template assets and live template configuration unchanged. The draft does not yet generate PDFs.

Run `preview/server.cjs` with `PORTAL_SOURCE_DIR` pointing at the original checkout. It reads a snapshot and binds only to 127.0.0.1:5185. It provides no write endpoints. No operational data or credentials are committed with the draft.

The Import portal backup button accepts the existing backup export format (`data.boardStore` and `data.installers`). Imports stay in browser memory and disappear on reload. Do not commit backup exports.

Morning Meeting, Filtering Board, Materials and Mustang are excluded from the rebuild.
