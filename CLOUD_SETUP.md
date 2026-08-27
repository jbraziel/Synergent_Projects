# Cloud Persistence Setup (Development / Option B)

The app now supports two storage modes without changing the UI:

- **Local mode:** SQLite + local files (automatic fallback; good for local testing)
- **Cloud mode:** Supabase database + private Supabase Storage (recommended while the app remains on Streamlit Community Cloud)

The code intentionally isolates persistence so a later move to an internal server can replace Supabase with SQL Server + a network share without rewriting the proposal UI.

## 1. Create a Supabase project

Create a project for the Marketing Proposal Generator. Do not store member-level extracts or other sensitive member data in this development project.

## 2. Create the database tables and private file bucket

In Supabase, open **SQL Editor**, paste the contents of `supabase_schema.sql`, and run it once.

The application seeds its default pricing settings automatically the first time it connects.

## 3. Add the Python dependency

`requirements.txt` already includes `supabase==2.31.0`.

## 4. Configure Streamlit Secrets

In Streamlit Community Cloud, open the app settings and add the values shown in `secrets.example.toml`.

Do **not** put credentials in GitHub. The app reads them from Streamlit Secrets.

For this server-side Streamlit app and private bucket, use a server-side Supabase key that has the required database and storage permissions. Keep that key only in Streamlit Secrets.

## 5. Reboot the Streamlit app

Once the secrets are present, the app automatically switches to cloud mode. You can verify it under:

**Admin > Storage Status**

You should see:

- Proposal Database: **Supabase Cloud**
- Proposal Files: **Supabase Storage**

## 6. Test persistence

Create a test proposal, generate it, and verify both downloads work. Then reboot the Streamlit app and confirm:

1. The proposal still appears in Proposal Library.
2. Edit reopens all of its inputs.
3. Draft PPTX still downloads.
4. Pricing CSV still downloads.
5. Admin pricing history is still present.

## Pricing behavior

Each NEW proposal freezes a copy of the current Admin pricing schedule when the proposal is started. Later Admin price changes do not recalculate that proposal.

A duplicated proposal is considered a NEW proposal. It copies the proposal configuration and manual proposal-price overrides, but starts with the current Admin rate schedule.

Every time a proposal is generated, the app also writes a separate immutable pricing audit snapshot to `proposal_pricing_snapshots`.

## Admin access during development

The default Admin users are:

- Jen Braziel
- Melanie Moore

This is UI-level access based on the selected user name, not true authentication. When the app moves to the internal Synergent server, Admin authorization should be tied to the actual signed-in employee identity.

## PDF conversion note

The current PowerPoint-to-PDF feature uses Microsoft PowerPoint automation and therefore requires Windows + installed PowerPoint. It is intentionally unavailable while files are running in cloud mode. This can be restored when the app moves to the internal Windows server.
