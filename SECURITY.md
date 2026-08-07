# Security and privacy

## Supported use

This portfolio code is intended for private Google Workspace deployments. It must not be connected to operational procurement data through a public or anonymous web-app deployment.

## Required controls

- Deploy as **User accessing the web app**.
- Restrict deployment access to the intended Google account or Workspace organization.
- Configure `ALLOWED_DOMAIN` or `ALLOWED_EMAILS` in Script Properties.
- Store spreadsheet, template, and Drive folder IDs outside version control.
- Grant the source spreadsheet and archive folder only to authorized users.
- Review Apps Script execution logs and Drive sharing periodically.

## Reporting a vulnerability

Open a private GitHub security advisory for vulnerabilities. Do not include real employee, vendor, banking, budget, letter, spreadsheet, document, or Drive data in an issue.
