# Copilot instructions — BookingsWithMeSignatureManager

> Canonical standards live in the `dev-standards` repo on SOUNDWAVE/Gitea.

## What this repo is

A **Microsoft 365 / Bookings-with-me** utility: a PowerShell helper plus a
packaged solution (`.zip`). Not a Home Assistant component.

## Repo shape

- `Get_MailboxGUID_from_user_for_BookWithMe.ps1` — PowerShell helper.
- `BookwithMesignaturemanager_*.zip` — packaged solution (binary; not unpacked
  into version control — agents can't read inside it).

## Conventions

- M365/PowerShell tooling: no HA component scaffolding applies.
- The core solution is zipped — to work on it meaningfully, unpack into the repo
  first.
- Never commit tenant IDs, mailbox GUIDs for real users, or credentials.

## Never

- Don't commit secrets or real tenant/user identifiers.
