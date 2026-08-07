# Security and privacy

## Supported use

This portfolio code is intended for controlled Google Workspace deployments. The URL may be reachable by signed-in Google accounts, but application data remains unavailable until allowlist and Gmail OTP checks pass.

## Required controls

- Deploy as **Execute as: Me** and **Who has access: Anyone with Google account**.
- Register only approved Gmail addresses in `DB_Pengguna`; protect Owner/Admin accounts.
- Monitor Gmail OTP quota and Apps Script execution logs.
- Store spreadsheet, template, and Drive folder IDs outside version control.
- Do not share the source spreadsheet or bound Apps Script project with application users.
- Keep generated documents in a controlled Shared Drive/folder and review membership periodically.
- Review Apps Script execution logs and Drive sharing periodically.

## Reporting a vulnerability

Open a private GitHub security advisory for vulnerabilities. Do not include real employee, vendor, banking, budget, letter, spreadsheet, document, or Drive data in an issue.
