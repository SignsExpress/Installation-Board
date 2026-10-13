# Standalone Vehicle Pricing — first working version

This is a separate application. It has its own server, login cookie, users, companies and pricing database. It does not load portal users/jobs or contact CoreBridge. The existing portal startup and module code remain unchanged.

## Local preview

- Build: `npm.cmd run build --prefix vehicle-pricing-product`
- Start: `node vehicle-pricing-product/server.cjs`
- Open: http://127.0.0.1:5192
- Owner credentials are generated on first startup in ignored `.runtime/owner-access.json`.
- The local demonstration company uses `demo` / `Preview-Only-5192!`. These are preview-only credentials, not a hosted account.

The owner can create companies, add administrators/estimators, reset company-user passwords and suspend access. Company administrators complete pricing setup and edit settings. Estimators use their company's saved settings.

## Pricing and setup

`baseline.json` contains pricing settings only, copied from Matt's local configured settings at initial extraction. There is no job/customer/account data. New companies get their own snapshot; existing custom settings are not overwritten when this file is updated.

The wizard asks for four reference prices for each existing vehicle size (Medium, Large, Luton). It uses vector illustrations and explicit assumed body areas of 30, 40 and 45 square metres. Illustrations are explanatory, not measured vehicle artwork. The formula functions in `src/formula.mjs` were extracted from the existing calculator. Calibration fits one price scale per vehicle size, reports actual calculated prices alongside targets, and retains the detailed Settings/Training Bank for finer adjustments. It does not promise an exact fit for arbitrary prices.

## Validation

Product and portal builds pass. Automated tests cover salted passwords, cross-company read/write isolation, inaccessible portal APIs, suspended-session rejection and calibration. Browser QA covers demo login, all three setup steps, review/save and the extracted vehicle drawing calculator, with no console errors.

## Hosting next

This first version runs locally. Before external companies use it, create a separate hosting service and URL, set NODE_ENV=production and PRICING_HOST=0.0.0.0, provision persistent storage and backups, and point PRICING_DATA_DIR there. Supply PRICING_ADMIN_USERNAME/PRICING_ADMIN_PASSWORD securely for first initialization. Sessions are memory-only and users must sign in after a restart; use one service instance with the file-backed store. No production service was created by this change. A package lock and standalone install should be generated on the hosting setup pass; local development currently resolves the existing workspace dependencies.
