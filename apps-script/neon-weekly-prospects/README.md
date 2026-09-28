# Neon CRM → Google Sheet weekly prospects

Every Monday at 8am it adds 10 new people to the **Prospects** tab. Each person has opened our emails, has never donated, and has both a Windfall ID and a Net Worth in Neon. Anyone pulled before is skipped. Notes typed in the **Notes** column are posted to that person's Neon account as an Activity, and the row is marked ✅ in **Synced to Neon**.

## Setup (about 10 minutes)
1. Create a Google Sheet, then go to **Extensions → Apps Script** and paste in `Code.gs`.
2. Open **Project Settings → Script Properties** and add `NEON_ORG_ID` and `NEON_API_KEY`. Create the key in Neon under Settings → User Management → API access.
3. Run `logSearchFields` once and look in the Execution log for:
   - the custom field IDs for **Windfall ID** and **Net Worth**. Put these in `WINDFALL_ID_FIELD` and `NET_WORTH_FIELD`.
   - the exact names of the email-opened and donation-count search fields. Change `EMAIL_OPENED_FIELD` and `DONATION_COUNT_FIELD` if yours are named differently.
4. Run `setup` and approve the permissions. This creates the Monday trigger and the notes trigger.
5. Optional: use **Neon → Pull 10 prospects now** to test.
