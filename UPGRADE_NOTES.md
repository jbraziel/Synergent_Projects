# Marketing Proposal Generator - Upgrade Notes

## Included in this version

- Sidebar proposal workflow with persistent completion progress.
- Progress bar, section count/percentage, and "Ready to Generate" state.
- Section completion checkbox moved into the sidebar with a Continue action.
- Library and New Proposal actions available directly from the sidebar.
- True New Proposal reset so values from the prior proposal do not carry forward.
- Proposal duplication from the Proposal Library.
  - Creates a separate Draft proposal.
  - Copies proposal inputs, targeting, components, pricing assumptions, custom costs, and manual pricing overrides.
  - Resets generated/sent/signed file history and completion flags.
  - Stores `copied_from_proposal_id` lineage.
- Proposal ID and proposal type displayed in the library.
- Copied proposals show their source proposal in the workspace.
- Missing saved fields added, including EMP subscriber count and campaign pricing overrides.
- Pricing detail CSV is now generated when the proposal is generated, rather than only when marked sent.
- Pricing export now includes:
  - proposal metadata and source proposal ID
  - targets and campaign components
  - hourly quantities and rates
  - raw list/print costs and 35% markup
  - email labor and send counts
  - each straight cost and whether it was included
  - each custom cost individually
  - one-time and repeating cost totals
  - calculated pricing and final/manual proposal pricing
  - EMP tier pricing details
  - Credit Card campaign inputs/goals
- Pricing detail remains available as a download after draft generation.
- Marking a proposal sent reuses the pricing record created with that draft instead of silently recalculating a different one.
- SQLite database and generated-file locations can now be configured with environment variables:
  - `PROPOSAL_DB_PATH`
  - `PROPOSAL_OUTPUT_ROOT`
- Duplicate `streamlit` entry removed from `requirements.txt`.

## Persistence still requires a permanent host/storage location

The code is now ready to point SQLite and generated files to a persistent mounted folder through the two environment variables above. If the app is running on a host whose local filesystem is ephemeral (such as Streamlit Community Cloud), setting a different local path on that same ephemeral filesystem will not make it permanent.

For long-term production use, use one of these approaches:

1. Host the Streamlit app on an internal server/VM with a persistent network or local volume, and point `PROPOSAL_DB_PATH` and `PROPOSAL_OUTPUT_ROOT` to that volume.
2. Migrate proposal records to a hosted PostgreSQL database and generated PowerPoint/PDF/CSV files to persistent object/file storage.

The current upgrade deliberately does not hard-code external database credentials or a storage provider.

## Validation performed

- `app.py`, `database.py`, and `generate_proposal.py` compile successfully with Python.
- All three files parse successfully as Python AST.
- SQLite schema migration, save/load behavior, and `copied_from_proposal_id` persistence were tested against a temporary database.
- A full Streamlit browser launch was not possible in the build environment because Streamlit itself is not installed there; test the interactive UI after deploying/installing the requirements in the normal app environment.

## UI refinement – compact library controls and full names
- Moved the active section's **Mark section complete** checkbox to the top of the proposal sidebar, before proposal details/progress, and styled it as a more visible completion card.
- Reworked Proposal Library controls into compact, aligned groups:
  - Actions are now one row: **Edit / Copy / Delete**.
  - Files are now one row of four equal-size icons: Draft / Sent / Signed / Pricing.
  - Missing files show disabled placeholders so every row stays aligned.
- Removed Doug from the user selector and added Melanie Moore.
- User selector now uses full names: **Jen Braziel, Shannan Heacock, Erica Vachon, Melanie Moore**.
- MSR display/filter uses **Shannan Heacock** and **Erica Vachon**, while remaining compatible with older proposals saved as `Shannan` / `Erica`.
- Existing saved `updated_by` and lock names using legacy first names are normalized for display.

## EMP review/generation isolation fix
- The Generate Proposal review now shows EMP-specific subscriber/tier/monthly pricing instead of campaign targets and auto-loan campaign costs.
- EMP generation no longer requires campaign targets or campaign components.
- EMP-only PowerPoint placeholder values are cleared before every generation run, preventing an EMP proposal from causing a later campaign proposal to be treated as EMP due to module-level state.
- Campaign target/component collections are empty for EMP generation, preventing stale campaign selections from leaking into EMP output.

## 2026-08-27 — Admin Pricing + Cloud Persistence Architecture

### Admin pricing area
- Added an Admin navigation area (default Admin users: Jen Braziel and Melanie Moore; configurable with `ADMIN_USERS`).
- Moved hard-coded labor rates, markups, email-send fee, fixed costs, four-campaign discount, EMP tier prices, EMP add-ons, and EMP implementation fees into database-backed pricing settings.
- Admins can edit rates and activate/deactivate fixed costs.
- Admins can mark fixed costs as repeating per campaign.
- Admins can add new fixed-cost items without editing Python.
- Added pricing-change history with user, timestamp, old/new value, repeat behavior and active/deactivated changes.

### Historical pricing protection
- Every new proposal freezes the current pricing schedule into `pricing_settings_snapshot`.
- Admin changes therefore apply only to new proposals, not existing ones.
- Legacy proposals created before this feature are frozen to the legacy hard-coded rate schedule when opened.
- Duplicated proposals are treated as new proposals and start with the current Admin rate schedule.
- Every generation writes an immutable pricing audit record to `proposal_pricing_snapshots`.

### Cloud persistence (Option B)
- Added automatic local/cloud database switching.
- With no cloud secrets, the app continues to use SQLite and local files.
- With Supabase secrets, proposals/pricing use Supabase and generated proposal artifacts use private Supabase Storage.
- Added `supabase_schema.sql`, `CLOUD_SETUP.md`, `secrets.example.toml`, and `.gitignore` protections.
- Generated PPTX filenames now include a time-based version suffix so same-day generations do not overwrite each other.
- Sent and Signed snapshots now work through the storage abstraction rather than assuming a local disk.

### Future Option A
- Persistence is isolated behind `database.py` and `file_storage.py`, so the later internal migration can replace Supabase with SQL Server and the internal network share without redesigning the proposal screens.


## Compact timestamp display
- Proposal Library and Admin Pricing History timestamps now display in Eastern Time as `M/D/YYYY h:mm AM/PM` (for example, `8/27/2026 6:36 PM`).
- Supabase timestamps remain stored in UTC; this is a display-only conversion.
