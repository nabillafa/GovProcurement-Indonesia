# Security review summary

Review date: 7 August 2026  
Scope: `PPBJ_Automation.gs`, `PPBJ_Web.gs`, `Index.html`, manifest, documentation, checks, and workflow files.

## Threat model

The protected assets are procurement records, employee/vendor identity data, budget details, generated documents, and Google Drive links. The main entry points are the Apps Script web-app URL and callable `google.script.run` server functions. Trust boundaries exist between the browser, Apps Script runtime, Google Sheets, Google Docs, and Google Drive.

The relevant attacker is a person who discovers the deployment URL but is not authorized to view or change organizational data. A second risk is accidental disclosure through source control.

## Controls in this portfolio release

- All real Google asset IDs and operational demo records were removed.
- Organization-specific values are loaded from Script Properties.
- The web app executes as Owner so users do not inherit spreadsheet or bound-script edit access.
- A short-lived six-digit code is sent only to an active `DB_Pengguna` Gmail address.
- The authenticated email is cached against `Session.getTemporaryActiveUserKey()` for a bounded session.
- Server endpoints enforce role, per-package ownership, and per-account budget-import permission.
- The code no longer enables cross-origin framing.
- Browser links are restricted to HTTPS URLs on Google Docs and Google Drive hosts.
- Rich-text and formula hyperlinks are resolved server-side, while label-only or non-Google values fail closed.
- Dynamic package buttons use data attributes instead of interpolating IDs into JavaScript source.
- CI checks syntax and common sensitive-data patterns on every push and pull request.

## Residual operational requirements

Security still depends on deployment and Drive configuration. Deployers must select **Execute as: Me** and **Who has access: Anyone with Google account**; application access then fails closed until Gmail OTP verification succeeds. Shared Drive membership, Owner/Admin accounts, mail quota, and generated-file permissions require operational review. The portfolio repository does not include operational spreadsheets or templates, so end-to-end Google authorization was not dynamically tested.

## Result

No known secret, real personal record, office-specific identifier, or Google asset ID remains in the published code scope. The local syntax and sensitive-data checks pass.
