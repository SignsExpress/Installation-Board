# Portal rebuild

Work stays on `codex/portal-rebuild`. Do not merge or deploy to production until the user authorizes the switch. The original checkout and operational data must remain untouched.

## Working version

Build, then run `npm.cmd run preview:rebuild` with `PORTAL_SOURCE_DIR` pointing at the original checkout. The rebuild binds to `127.0.0.1:5186` and uses the existing React modules and Express API. The visual mockup remains on port 5185.

The runner creates a separate database under ignored `outputs/rebuild-runtime`. A random test password for the copied Matt profile is recorded in ignored `preview-access.json`; the original password is unchanged. Never commit this file or runtime data.

On first initialization, `PORTAL_PREVIEW_BACKUP` may identify a full portal backup, or `PORTAL_PREVIEW_SNAPSHOT` may identify captured installation/design data. Existing runtime files are preserved on restart. Do not overwrite test work for re-import without reviewing it.

SMTP, push settings, TimeMoto credentials, live CoreBridge synchronisation and queued social posts are disconnected in this runner. Sending/publishing routes return an error. CoreBridge and OpenAI credentials are currently absent, so job pulls and AI generation remain unverified.

## Design and scope

Original Signs Express assets supply purple `#5f3c74` and teal `#0f98a5`. The new stylesheet applies the Evo-inspired shell under `@media screen` only.

Morning Meeting, Filtering Board, Materials and Mustang launch points are removed. Their local routes redirect to Home. Remaining modules use their established implementations, including WIP, Order Panels, IGLOO and Permissions, which were omitted from the initial static mockup.

## Pro-forma preservation

Keep the established generator, editor, calculations, print styles, template assets and live template configuration unchanged. The generator and original stylesheet were compared with the original checkout and match, ignoring line-ending differences. Custom live template data/assets still need copying before confirming exact live PDF output.

## Validation

Build and server syntax checks passed. Twelve authenticated module read APIs passed. Installation create/read/delete, design task create/schedule/delete and WIP persistent-save checks passed against test data. Browser review confirmed the branded home page, WIP controls and pro-forma editor.
