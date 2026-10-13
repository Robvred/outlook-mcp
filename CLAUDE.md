# Repository rules

This repository is public. It contains generic tooling only. Nothing that identifies the
maintainer's employer, colleagues, vendors, projects or mailbox content may ever be committed.

## Anonymization (mandatory)

Applies to code, comments, docstrings, tests, fixtures, documentation, commit messages,
branch names, pull request titles and descriptions, and issue text.

- No real organization names: employer, subsidiaries, customers, vendors, partners.
- No real people: names, initials, display names, roles tied to a real person.
- No real email addresses, domains, mailbox names, folder paths, server names, URLs,
  SharePoint sites, tenant IDs or internal system names.
- No real project, application, programme or contract names, ticket or order numbers,
  amounts, dates or figures taken from real work.
- No content from real emails, attachments, meetings or documents, even partial, even
  paraphrased, even "just as an example".
- Use neutral placeholders only, consistently:
  - People: `DOE John`, `Jane Smith`
  - Addresses and domains: `user@company.com`, `example.com`
  - Folders: `user@company.com/Inbox/ProjectA`, `user@company.com/Projects/ProjectA`
  - Organizations and vendors: `Company`, `Vendor A`, `Vendor B`
  - Projects and systems: `Project A`, `System X`
  - Amounts and identifiers: obviously fake values (`12345`, `INV-0001`, `1 000 EUR`)
- When describing a bug found on a real mailbox, describe the mechanism only
  (Outlook/COM behaviour, filter, timing), never the data that revealed it.
  Write "on a real Exchange mailbox", not who, what or which folder.

## Data that must never be committed

- Mailbox caches, exports and indexes: `*.msg`, `*.eml`, `*.pst`, `*.ost`, `*.sqlite`, `*.db`,
  `email_cache.json`.
- Logs and audit trails: `audit.log`, `*.log`.
- Saved attachments, screenshots of a real mailbox, copied email bodies.
- Credentials, tokens, client secrets, connection strings.

`.gitignore` covers the usual file types; it is a safety net, not a substitute for review.

## Before every commit or pull request

1. Search the diff for real names, domains, amounts and identifiers:
   `git diff --cached | grep -iE "<terms you know are real>"`.
2. Re-read docstring and help-text examples: they are the most common leak.
3. Re-read the PR title and description with the same rules.
4. If something real was pushed, stop, replace it with a placeholder in a new commit, and tell
   the maintainer: history and PR pages keep the old text and may need a separate clean-up.

## Safety of the tools

- Outbound tools only create drafts. Never add a call that sends mail (`.Send()`), responds to
  meetings, or creates calendar invitations.
- Destructive operations stay recoverable: delete means move to Deleted Items; file writes never
  overwrite existing files.
- Every write operation is logged to the local audit log, which stays on the user's machine.
