# Credit Builder Flow — Follow-Up Gmail Draft Generator (Google Apps Script)

Automates post-meeting follow-ups for `Cold Call Broker List` (`Booked Meetings` tab) whenever Google Meet saves a **Notes by Gemini** Google Doc in `creditbuilderflow@gmail.com`'s Google Drive.

## What It Does
1. Searches Google Drive for recent Google Docs (`Notes by Gemini` / `Recover Lost Revenue`) not yet recorded in the sheet.
2. Reads **all document tabs** (`Full notes`, `Quick notes`, `Transcript`) via `DocumentApp.openById()` and extracts any linked Google Calendar `?eid=` event ID.
3. Matches each meeting doc to the single corresponding row in **`Booked Meetings`**:
   - Primary match: Calendar Event ID (`?eid=` hyperlink decoded against Column O `Calendar Event ID`).
   - Fallback match: Dealership name parsed from Doc title + meeting date (`Africa/Johannesburg`).
4. Sends the meeting notes & transcript to **Gemini API** (`gemini-2.5-flash` by default) to extract ONLY explicitly stated parameters (`dpName`, `dpEmail`, `declines`, `marginK`, `banks`, `concern`). Never invents or guesses numbers.
5. Builds the Option A1 personalised URL for `https://dealerships.volvemedia.com/` omitting missing optional params.
6. Creates a **Gmail Draft** (never sends automatically) to the F&I contact (and DP if `dpEmail` was captured).
7. Appends non-destructive tracking columns **after Column Y** (`Follow-Up Status`, `Portal URL`, `Gmail Draft ID`, `Notes Doc ID`, `Extracted JSON`, `Processed At`). Rows missing `declines` or an email are flagged **`Needs review`**.

---

## Setup Instructions
1. Open the **Cold Call Broker List** spreadsheet (`1BkZafvLWWWrKm56r_c8IThjIEXcPOZHWnEF4unA-K5g`).
2. Click **Extensions → Apps Script**.
3. Open **Project Settings** and enable **"Show appsscript.json manifest file in editor"**.
4. Replace `appsscript.json` with [appsscript.json](./appsscript.json).
5. Create or replace `Code.gs` / `FollowUpDraftGenerator.gs` with [FollowUpDraftGenerator.gs](./FollowUpDraftGenerator.gs).
6. In **Project Settings → Script Properties**, click **Add script property**:
   - **Property:** `GEMINI_API_KEY` → **Value:** your Google AI Studio Gemini API key.
   - *(Optional)* **Property:** `PORTAL_BASE_URL` → defaults to `https://dealerships.volvemedia.com/`
   - *(Optional)* **Property:** `GEMINI_MODEL` → defaults to `gemini-2.5-flash`
   - *(Optional)* **Property:** `LOOKBACK_DAYS` → defaults to `7`
7. Select **`testDraftGeneratorDryRun`** in the top toolbar and click **Run** (accept OAuth permissions when prompted; this dry run does not edit Drive, Sheets, or Gmail).
8. Select **`installTenMinuteTrigger`** and click **Run** once to schedule **`processCompletedMeetings`** every 10 minutes.

## Notes & Safeguards
- Columns **A–Y** in `Booked Meetings` are strictly read-only and never overwritten.
- No email is ever sent automatically; every output lands in **Gmail Drafts** for manual review and clicking **Send**.
- If a Gemini Notes doc is re-edited and you want to regenerate the draft, clear `Notes Doc ID`, `Gmail Draft ID`, and `Follow-Up Status` for that row after deleting the old draft.
