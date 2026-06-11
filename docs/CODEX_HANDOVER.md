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
- The report now includes an `Order materials` workflow that lets users tick/untick the material list, finalise it, group like materials, estimate panel sheet usage, and suggest panel offcuts from waste. Vinyls and laminates stay as a clean grouped list without a roll-usage calculator.
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
