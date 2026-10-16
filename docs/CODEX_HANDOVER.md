# Codex Handover

Last updated: 11 June 2026

This note is for continuing development from another computer or a fresh Codex window.

## Resume Instructions

1. Open the Installation Board repository.
2. Pull the latest `main` branch from `origin`.
3. Codex should automatically read the repository-level `AGENTS.md`.
4. Read this file before changing the application.
5. Continue the user's newest request, then build, commit, and push automatically.

## User Workflow Preference

The user wants Codex to handle the full workflow:

- inspect the existing implementation
- make the requested change
- run the production build
- run relevant checks
- commit only the relevant application files
- push automatically

Do not ask the user to run PowerShell for normal development tasks.

## Repository And Deployment

- Repository: `https://github.com/SignsExpress/Installation-Board.git`
- Normal branch: `main`
- Production is deployed through Render from pushed repository changes.
- Frontend: React/Vite.
- Backend: Express in `server/index.js`.
- Production build command: `npm.cmd run build`.
- Server syntax check: `node --check server/index.js`.

## Current Product State

The portal contains multiple operational modules including Installation Board, Morning Meeting, Design Board, Filtering, Attendance, Holidays, Mileage, Materials, Vehicle Pricing, RAMS, Social Post, Description Pull, Pro-Forma, Subcontractors, and Notifications.

### Morning Meeting

The Morning Meeting Board is available at `/morning-meeting` and launches from the Installation Board.

- Shows yesterday's installation jobs, today's installation jobs, and artwork approved yesterday.
- Also shows tomorrow's installation jobs.
- Yesterday's installs can be selected and given job-specific Morning Meeting notes. Saving applies each note to every matching `EST`/`ORD`/`INV` instance on the Installation Board.
- Morning Meeting notes are stored separately from ordinary job notes and appear in a dedicated green note box on Installation Board cards.
- Typing `@` in a job-specific Morning Meeting note suggests portal users. Selected mentions use the full name and appear bold in the emailed meeting summary.
- Includes job references, customers, descriptions, values, installers/job types, addresses, contacts, and relevant notes.
- Meeting notes can be entered and emailed to every user with an Email value in their Permissions profile.
- Meeting emails use a cleaner card layout and place persisted job notes directly beneath the relevant jobs.
- Uses the server SMTP configuration already shared with credit applications.
- Can fetch and print a Job Materials Report for yesterday's approved artwork.
- The Job Materials Report reads the matching Estimate/Order item assemblies through `Items`, `Components`, `ChildComponents`, and `AssemblyDataJSON`.
- It separates finished-item quantity and size from stock size and material consumption. Never treat roll-material `TotalQuantity` as the number of finished items.
- It also lists nested materials, laminates, fixings, labour/service parts, time, and employee counts where CoreBridge supplies them.
- The visible report is deliberately compact: one `quantity + finished size` heading, a deduplicated checklist of named production materials/fixings, and one total estimated production-time line converted to hours and minutes. Do not expose generic assembly wrappers or machine/layout internals.
- The report uses the existing server-side OpenAI configuration to organise the full grounded CoreBridge evidence into concise production overviews. AI must not invent details; when unavailable, the rule-based summary remains as fallback.
- Job Materials Report item summaries display in a four-column desktop grid beneath each job reference.
- The report toolbar now shows a live loading state while the background fetch runs, with steps such as CoreBridge lookup and AI simplification.
- Material names in the visible checklist should have trailing stock sizes stripped out so production sees the clean material name first.
- The report now includes an `Order materials` workflow focused on panel stock only. It groups returned panel materials under the actual material name, then renders sheet-layout diagrams showing how the finished panels fit inside the chosen sheet size.
- The host landing page now includes a `Job Materials` launch module that opens the Morning Meeting materials workflow directly.
- The panel ordering strip now uses the tidied material list as its source, gives each material its own suggested sheet size, and allows per-material overrides including custom sizes.
- Latest related commits:
  - `01e86dd Pull materials from CoreBridge works orders`
  - `7fb0d3f Add approved job materials stock sheet`
  - `7f8cc91 Email morning meeting notes`
  - `7c0bec6 Add morning meeting outline`

### Installation Board

- The rolling installation history and full card layout were restored.
- Installation values were corrected, including duplicate-total handling and invoice-value prioritisation.
- Mobile cards now show installation notes.
- Latest related commits include `db6942b`, `47d718e`, `e2ff130`, `5cb3388`, `8b0635f`, and `d1ca4ad`.

### Design Board

The Design Board has recently been substantially redesigned:

- Days run vertically down the board.
- New Orders, Unallocated, Awaiting Sign-Off, day sections, and Order with Salesperson use a shared card layout.
- Cards show added/amended/chased timing.
- Cards can be marked priority, chased, edited, deleted, and given a red designer note.
- Cards open a copyable detail popup.
- Uploader photo/initials appear beside the card identity for newly pulled or refreshed cards.
- The toolbar has approval totals and configurable daily, weekly, and monthly targets.
- The Pull control has an integrated loading spinner.
- Dragging a card near the top or bottom edge automatically scrolls the board.
- Latest related commit: `c93ed17 Add design board drag auto-scroll`.

### Filtering Board

- Filtering now uses the same cards and layout as the Design Board.
- All approved jobs appear in one `Approved Jobs` section.
- Cards display `Date Approved`.
- Clicking a Filtering card opens the matching detail/copy popup.
- Latest related commit: `2267b4c Align filtering board with design board`.

### Vehicle Pricing

- Added `Luton Van / Vauxhall Movano Luton`.
- Source tyre diameter was `740.98mm`; the outline is intended to be at 10%.
- Added vertical spacing so the outline has top and bottom breathing room.

### Other Recent Work

- Added an admin-only CoreBridge Explorer at `/corebridge-explorer`. It groups known CoreBridge APIs by category, filters them by search, runs a selected API against a job reference/search value, and shows the untouched response plus attempted paths for mapping work. Credentials remain server-side.
- Mileage rate was changed to `GBP 0.55` per mile.
- RAMS first-aid facilities gained support for a custom hospital when postcode lookup is incomplete.
- Design approval targets and progress bars were added.
- Design cards gained notes, contact numbers, net totals, priority state, and improved layout.
- Notifications are grouped under day headings, newest day first, with a two-column desktop feed.

## Important Known Issue To Preserve

CoreBridge line item mapping has previously been incorrect or blank on Design Board cards.

The required behaviour is:

- Line item name must come directly from CoreBridge's `Line Item Name`.
- Category must come directly from CoreBridge's `Category`.
- Do not infer these from descriptions, product codes, or job types.
- Do not add per-job or hard-coded overrides.

Examples previously supplied by the user:

- `EST-3644`
  - Line item name: `New Entrance Sign Panel`
  - Category: `Panels`
  - Second line item name: `Installation`
  - Second category: `Installation`

If this is revisited, inspect the raw CoreBridge response and map the exact fields before changing presentation logic.

## Working Copy Safety

The office working copy commonly contains unrelated local changes and OneDrive conflict files. Do not stage or alter them.

Known local-only/unrelated items at the time of this handover:

- modified `data/jobs.json`
- modified `data/users.json`
- `data/backups/`
- `server/auth-store-LAPTOP-MF2L0L1V.js`
- `server/index-LAPTOP-MF2L0L1V.js`
- `src/App-LAPTOP-MF2L0L1V.jsx`
- `src/index-LAPTOP-MF2L0L1V.css`
- `tmp_pdf_compare/`
- `tmp_pdf_inspect/`

Always stage explicit paths, for example:

```powershell
git add src/App.jsx src/index.css
git commit -m "Describe the change"
git push
```

## Recent Commit History

- `01e86dd` Pull materials from CoreBridge works orders
- `7fb0d3f` Add approved job materials stock sheet
- `7f8cc91` Email morning meeting notes
- `7c0bec6` Add morning meeting outline
- `db6942b` Show installation notes on mobile cards
- `47d718e` Resolve missing installation invoice values
- `e2ff130` Prioritize invoice values by branch
- `5cb3388` Fix installation values and duplicate totals
- `8b0635f` Restore rolling installation board view
- `d1ca4ad` Restore full installation history and card layout
- `c93ed17` Add design board drag auto-scroll
- `2267b4c` Align filtering board with design board
- `fba70f1` Stabilise design pull loading button
- `d6d21ea` Add design order pull loading animation
- `ceee2fc` Modernise design board header and uploader
- `f9fea74` Add design approval progress dashboard
- `b074f39` Tidy design card contact and date
- `f174e01` Rethink design board card layout
- `44a293a` Add design card notes and polish styling
- `c368891` Add order with salesperson design lane
- `8e97f66` Add vertical spacing to Luton outline
- `a12546c` Add Vauxhall Movano Luton pricing outline

## Starting Locally On Another Computer

After cloning or pulling:

```powershell
npm.cmd install
npm.cmd run dev
```

The local frontend normally opens at:

```text
http://localhost:5173
```

The home computer must also have:

- access to the GitHub repository and working Git credentials
- Node.js/npm installed
- any private environment variables needed for CoreBridge or other external services

The auto-build/commit/push preference is stored in `AGENTS.md`, but Codex command approvals and private credentials are machine-specific. A fresh Codex window should diagnose and complete any missing local setup rather than changing application code to work around it.

Do not copy office `data/jobs.json` or `data/users.json` changes into Git merely to make the home machine match. Production operational data is separate from source-code deployment.

## Optional new skin — 8 October 2026

Classic remains the default application (`src/App.jsx`). The new interface is opt-in through `?skin=new`, remembers the choice in this browser, and has a permanent Switch to Classic link. `?skin=classic` explicitly restores the original application. `src/ModernApp.jsx` is the isolated early-access interface; its CSS and feedback wrapper load only when selected. Both use the existing live API, authentication, permissions and operational data. No copied preview database or disconnected test runner is used in production.

Share https://www.sxpreston.com/?skin=new with early-access users. Changes made there affect real live jobs. Removed modules are hidden/redirected in the new interface; Classic retains its existing module access for compatibility. Invoice PDF generator and print-preview functions are identical across skins. Browser QA covers new/classic switching, route persistence and original styling using the isolated preview API. Focused tests cover default selection, explicit overrides, URL preservation and PDF parity.

The existing holiday seed merge now also retains WIP state, matching the tested rebuild fix. Do not commit local operational data. Future behaviour fixes affecting both skins must be applied to both App.jsx and ModernApp.jsx until the early-access interface becomes the default.

### WIP behaviour restored — 8 October 2026

The new skin uses one combined draggable WIP board. Status lanes are always visible together with unscheduled work and ten working days starting today, skipping weekends, matching Classic's getWipBoardDays rule. There is no separate status/schedule toggle or Monday-based week navigation. Status lanes remain sticky on desktop to accept jobs dragged from the date lanes. Existing upload memory, multi-day planning, assignees, count-value and completion review rules are retained. Date and drag/drop transition regression tests live in tests/wip-board.test.mjs.

## Original design restored — 13 October 2026

The user rejected the new design. main.jsx now loads only the established App.jsx and index.css. The early-access switch is removed, saved skin preferences are cleared, and old ?skin=new links also load the original portal. The unused new-skin source remains in Git history/source for reference but is not imported or shipped in the production build. Do not re-enable it without an explicit user request. Live operational data, the existing PDF layout and the WIP data-preservation fix are retained.

## Four modules retired — 13 October 2026

Morning Meeting, Filtering Board, Materials and Mustang are removed from Classic desktop/mobile navigation and host/client home launchers. The Installation Board Morning Meeting shortcut is removed. Old module URLs (including client/nested routes) resolve to Home at startup through retired-modules.mjs. This is a page-entry removal: saved data and shared APIs/helpers remain to protect Design Board approval workflows, panel tools and other module dependencies. No remaining module screen, PDF generator, CSS or server implementation was changed. Build and tests/retired-modules.test.mjs checks pass.

## Standalone Vehicle Pricing foundation — 13 October 2026

vehicle-pricing-product is a separate local application on port 5192 with its own companies, users, sessions and settings. The portal does not import or serve it. Owner admin supports company creation, user creation/password reset and suspension. A first-login illustrated setup wizard calibrates four reference jobs across Medium/Large/Luton using extracted formula functions, then opens the extracted original calculator. Runtime data and owner credentials are ignored. Company isolation/password/calibration tests and both builds pass. See vehicle-pricing-product/README.md for limits and the independent hosting step still needed before external access.

## GOS Vehicle Liveries — 15 October 2026

Customer portal `/gos` has its own shared-password session (default requested password GOS123, override GOS_CUSTOMER_PASSWORD). GOS access grants no access to portal APIs. The host Home launcher is beside IGLOO. The React entry selects this page directly, without mounting the portal App or importing the unfinished standalone Vehicle Pricing app.

server/gos/seed.json contains the 58 spreadsheet vehicles, division assignments and five livery categories imported from GOS_Vehicles.xlsx. The workbook has no embedded images. Vehicle and approved-livery photos are uploaded through the authenticated customer portal. Existing register data are never overwritten by the seed: the module uses its own register.json and photos directory beneath path.dirname(DATA_FILE)/gos-vehicle-liveries, with an immediate previous-file backup.

Customers can edit/add vehicles, assign livery/division, filter/search the fleet and add livery categories/images. Order livery and Order repair support existing/new registrations. Repair boxes use normalized image coordinates and export actual annotated JPEG attachments, with machine-readable box data. Submission stores an order snapshot, uses a request id to prevent duplicate sends, records SMTP failure honestly and retains order history. New registrations enter the register after successful email acceptance; an order does not automatically change an existing vehicle's recorded livery.

Live email requires GOS_ORDERS_TO; no inbox was assumed. The question asking the user for the recipient remains outstanding. Without it, ordering is visibly disabled. Mail uses the existing credit-application SMTP transport/from settings. Do not send live test orders without the user's instruction. Photos must be present before ordering so emails always have a real visual reference. No fabricated photos were deployed.

Validation: tests/gos.test.cjs covers GOS-only authentication, seed counts, stale-edit rejection, duplicate registrations, reference-image attachment, annotated repair attachments, idempotency and SMTP failure. Browser QA on the isolated local preview covers filters, vehicle editor, drawing a box and capture/visual inspection of the resulting attachment. No test email was sent. Test preview/runtime files are ignored under outputs. Other portal modules and the standalone Vehicle Pricing product are unchanged.

## GOS PDF catalogue and repairs — 15 October 2026

GOSVans.pdf is now the protected source catalogue under server/gos/catalogue. The unchanged 19-page original and private page snapshots are served only after GOS authentication. Corner text supplies van/model labels; the user explicitly said to ignore the graphics themselves for classification. Page ranges: 1–9 Standard GOS, 10–14 Mechanical Electrical, 15–16 Electrical, 17–19 Custom Livery. Do not infer a livery from colours/logos or fabricate missing artwork.

Vehicle previews and orders can select a PDF design. Matching considers category, exact model generation and wheelbase, supports Pickup/Dropside aliases, and leaves missing/ambiguous models for manual selection. Explicit choices are saved with vehicle edits. Existing records and uploaded photos are preserved. No original PDF or operational database is overwritten by catalogue deployment.

The Livery designs page supports further PDF uploads (22MB, at most 100 pages), reads text/rendered previews with the existing PDF.js dependency and requires confirmation of the category for every page. Original PDFs and previews persist beside the GOS database. Only the PDF upload route has a 40MB JSON limit; other portal limits are unchanged.

PDF-based repair orders submit normalized page coordinates. The new server/gos/pdf.cjs uses pdf-lib to copy the original selected page and draw vector red boxes, preserving original page dimensions/artwork. Emails attach a PDF plus box-coordinate JSON, not a raster replacement. Legacy image-based orders are retained for backward compatibility. Browser QA tested PDF import, register preview, repair-box drawing, local email capture and rendering of the resulting PDF. Eleven regression checks and the portal build pass. No live email was sent; GOS_ORDERS_TO still awaits the user's recipient address.

## GOS order simplification and live mail — 15 October 2026

GOS mail now defaults To eddy.d'antonio@signs-express.co.uk and Cc matt.rutlidge@signs-express.co.uk, as explicitly requested. It uses the exact existing customer-account/credit-application SMTP configuration and From address. GOS_ORDERS_TO/GOS_ORDERS_CC remain optional overrides. Contact name/email are not required or shown in the form. A protected, cached /api/gos/email-status check verifies the SMTP connection without sending a message.

Existing-registration orders use a searchable registration/driver/model result list. Vehicle, livery and PDF selection are read from the saved register by the server, including an optional vehicle-version check against stale forms. Livery/PDF selectors, PDF/source links and name/email fields are removed from that popup. Missing artwork must be assigned in the register first. New registrations retain a single visual livery/artwork chooser because no saved record exists for a new reg.

The supplied Gos logo.svg is in public/gos/logo.svg and used in the header/login. Colours were checked against the live gosheating.co.uk stylesheet: red #EE2F2C, orange #F7792A, teal #006D77, pale cream #FFF3EA and pale blue #BAE1E5. Styles stay scoped to GOS. Tests cover automatic selection, contact-free ordering, To/Cc including Eddy's apostrophe, diagnostic mail checks, PDF attachments and prior portal routes. Browser QA uses local capture only; no unsolicited test order email was sent.

## GOS simpler UI and replaceable PDFs - 15 October 2026

GOS now uses red/orange with neutral backgrounds throughout; teal, green success messages and pale blue accents were removed at the user's request. Vehicle cards, labels and order forms are more compact, with short instructions and prominent Order graphics / Order repair buttons. The graphics modal is narrower, its preview is smaller and the send controls stay visible. Repair drawing keeps the larger working area. All styling remains scoped to GOS.

The PDFs tab supports adding a new PDF or replacing an existing document, including the original bundled catalogue. Originals are retained under immutable filenames so saved orders/retries keep their original attachment. Replacements become the single active document, reuse page IDs only for unique matching label/category pairs (including reordered pages), and pin existing vehicle assignments. Removed/changed drawings require manual selection instead of guessing. Vehicle versions increment to reject orders from stale forms. A document revision prevents two uploads from overwriting each other. No operational data or unfinished Vehicle Pricing app is included in deployment.

Validation: production build, server syntax checks and 16 regression tests pass. Tests cover imported and bundled PDF replacement, reordered pages, missing drawings, retained order sources and stale replacement rejection. Local browser QA checks registration search, compact order form and PDF replacement selector. No live test order email was sent.

## GOS van deletion and individual PDF replacement - 15 October 2026

The vehicle editor now has Delete van followed by an in-page Are you sure confirmation with Cancel and Yes, delete van. The authenticated DELETE route requires the current vehicle version and archives the removed record in deletedVehicles rather than deleting its files or order snapshots. Archived entries are excluded from the public register. Later email retries cannot recreate the archived vehicle.

Replace PDF for this van is available directly in the same editor. It accepts a one-page PDF and saves the current vehicle edits and explicit PDF assignment together, with a version check. Each van has its own immutable source files and active document version. Vehicle-specific pages are excluded from automatic matching and selection for other vans and from the shared catalogue management. Other vans using the original shared page remain unchanged. Previous order snapshots retain their original attachment.

Validation: production build, server syntax checks and 18 regression tests pass, including deletion/auth/stale edits, per-van isolation, atomic field saves and retained historical PDFs. Initial local UI inspection verified the editor controls; a native browser-confirm test blocked its test tab, so implementation uses the in-page modal instead. No live van was deleted and no live test email was sent.

## Dashboard Customer Portals heading - 15 October 2026

Host Home now groups the existing IGLOO and GOS launch cards under Customer Portals instead of Tools. Existing permissions, destinations and new-tab behaviour are unchanged; the responsive column count includes the new section. Other module screens and operational data are untouched. Production build passes.

## GOS replacement PDF label matching fix - 15 October 2026

The first live catalogue replacement omitted the separator dot from the original corner labels. Exact string comparison generated new page IDs while keeping vans pinned to old IDs, leaving 49 links broken. Replacement matching and importer category lookup now ignore punctuation, spacing and letter case, while still requiring unique labels and the same livery category. Reordered pages retain their correct assignments; removed or genuinely changed designs still require review.

The live register repair uses the original catalogue label/category and uniquely matching current uploaded pages, with vehicle version checks. Before/after snapshots and the exact repair plan are retained locally under ignored outputs; no operational data is committed. The uploaded PDF is retained and historical order attachments are unchanged. Production build, server syntax checks and 19 regression checks pass, including a full punctuation-free and reordered 19-page replacement across all originally matched vans.

## Ireland Delivery calculator - 16 October 2026

Ireland Delivery is available to every signed-in portal user at /ireland-delivery, in desktop/mobile navigation and both dashboard Tools sections. It does not require another module permission or change saved user permissions. A user with all module permissions set to none can access the calculator and sees its dashboard launcher. Shared GOS/IGLOO customer-password sessions remain separate from portal login.

src/ireland/calculator.mjs reproduces Ireland-Import-Cost-Calculator.xlsx (Ireland calculator, C6:C20 and F6:G15): round EUR customs value first; duty on customs value; VAT on customs value plus duty; DHL tiers below EUR5 / EUR5-20 inclusive / above EUR20 at 2.5% minimum EUR14.50. Sum rounded EUR fees and convert each result back to GBP using the same saved FX rate. Default FX1.163467, duty2.7%, VAT23%. The UI starts with blank goods/delivery and zero insurance, provides a workbook example and reset, and exposes all assumptions in a collapsed section. All edits are calculator-only, not shared data writes.

EORI GB213176733000 and commodity 4911910090 are displayed as selectable reference numbers. Scope is GB to Republic of Ireland and the supplied DHL estimate; rates are editable, FX is not live, totals are before VAT recovery, and commodity does not automatically set duty. The workbook's invoice/waybill provenance and Irish Revenue guidance link are in calculation details. No workbook upload or operational-data changes are included in source deployment.

Validation: production build and six focused tests (workbook EUR/GBP parity, DHL boundaries, custom assumptions, invalid/missing/zero inputs and remaining module routes) pass. Workbook example gives EUR131.59 / GBP113.10 additional fees and GBP215.47 delivery plus charges. Local browser QA checks the example, reference numbers, all-permissions-none navigation and dashboard visibility. Other modules and the standalone Vehicle Pricing product remain unchanged.

## Ireland Delivery quotation wording - 16 October 2026

Added Quotation description below the reference numbers with the user's exact supplied title and paragraph. Copy description copies plain text and, where supported, HTML with the title bold for pasting into quotations. Clipboard errors fall back to selecting the visible description for manual copying; feedback is announced in a status region. Wording stays fixed as supplied rather than changing when calculator rates are edited. Production build passes. Calculator logic, operational data and other modules are unchanged.

## Ireland Delivery compact layout and delivery total - 16 October 2026

Reduced module typography, input heights, spacing, results card, reference numbers and quotation section padding. Removed Load spreadsheet example while retaining Reset. The headline is now Total cost to Signs Express - excluding goods and uses deliveryTotal (courier delivery + insurance + import fees) in GBP and EUR. Additional import charges remain separately listed; goods-inclusive total remains secondary. Quotation wording/copy and calculator formulas are unchanged. Production build and four calculator tests pass. Local browser QA verified the headline GBP215.47/EUR250.69 for goods GBP280 and delivery GBP102.37, and inspected the compact layout.
