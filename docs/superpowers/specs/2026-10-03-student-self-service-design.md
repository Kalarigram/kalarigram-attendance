# Student self-service attendance — design

Date: 2026-10-03 · Status: approved by Madhan

## Goal
Replace the shared-phone kiosk with a page students open on their own phones to
register themselves and mark their own attendance.

## Decisions
| Topic | Decision |
|---|---|
| Identity | Device remembers the student (`kg_me`) after registering or picking their name once. "Not you? Switch" resets it. |
| New sign-ups | Active immediately. Server de-dups on phone (last 10 digits) and returns the existing ID + name. |
| Kiosk mode | Removed. PIN screen removed (PIN still sent if stored, server check unchanged). |
| Marking rules | Any enrolled class, once per class per day. Server also skips a second row for same date + class + student. |
| Enrollment changes | Admin only, in the sheet. No new API action. |

## Front end (`index.html`, single file, no build)
Views:
1. **Welcome** — shown when no `kg_me`. "Register" / "Find my name".
2. **Register** — name*, phone*, email, emergency contact*, class checkboxes* (active classes from cached `action=all`). Client validates phone (10 digits; `+91`/leading `0` stripped). POST `register`; on success show the Student ID and remember the student. Needs network; failures show a retry message (not queued).
3. **Find me** — search the union of all enrolled students by name or ID; tap to remember.
4. **Home** — name + ID, today's date, one card per enrolled class with a lamp and a "Mark present" button. Marked cards show a lit lamp and "Present today".

Kept patterns: single `action=all` load cached in `kg_data`; optimistic marks into `kg_queue`, flushed every 15 s / on reconnect / on focus; per-day `kg_marked`.

Freshly registered students are inserted into the local cache immediately, and `kg_me.classIds` is used as a fallback until the server list includes them.

Visual direction: lime-plaster ground, kalari red `#7a2e12`, brass `#b8862b`, banana-leaf green for "present"; Anek Latin; brass nilavilakku lamp per class that lights when marked (the one animated moment, disabled under reduced motion); poothara line drawing on the welcome view.

## Back end (`Code.gs`)
- `register`: normalise phone → last 10 digits; if an existing Students row matches, return `{ok:true, id, name, existing:true}` without writing.
- `attendance`: in addition to entryId de-dup, skip a record when a row with the same Asia/Kolkata date + Class ID + Student ID exists; report it as saved.
- No GET endpoint exposes phone/email/emergency contact.

## Testing
Manual run in a local browser against the live API: first visit → register → mark present → reload → switch → find me → offline mark → reconnect. Test rows are listed for deletion afterwards. Backend changes are deployed by Madhan as a new version of the existing deployment (same URL).
