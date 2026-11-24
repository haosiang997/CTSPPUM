---
name: "ctsppum-patient-notifier"
description: "Open the CTSPPUM Patient Notifier, a claude.ai Artifact that turns the CTSPPUM patient list Excel into ready-to-send WhatsApp messages (standard admission or change-of-date script, BM / EN / 中文, per-patient language and hand edits). Trigger when the user says \"open the notifier\", \"notify the patients\", \"send the WhatsApp reminders\", \"message the patients about their op date\", \"tell patients their surgery was postponed\", or has just produced patient_list.xlsx with the ctsppum-patient-list skill. Also use it to update or republish the notifier page itself."
---

# CTSPPUM Patient Notifier (Artifact)

The notifier is a single HTML page published as a claude.ai Artifact:

**https://claude.ai/artifact/PN4YWSqYjDKnqUJcgBekuz**

Its source is `assets/ctsppum-notifier.html` in this skill. The page reads the Excel file in
the viewer's browser only (SheetJS from cdnjs); no patient data is uploaded, stored or sent to
Claude. Each Send button is a link to `web.whatsapp.com/send?phone=…&text=…` that opens in a new
tab of whichever browser shows the page. That browser must be **Google Chrome**, where the user
is signed in to WhatsApp Web: opened in the Claude app or another browser, the links land in the
system default browser, which usually is not signed in.

## Workflow

1. **Get the list.** If the user has no `patient_list.xlsx` yet, run the
   `ctsppum-patient-list` skill first. Columns: Name · RN · Operation Planned ·
   Date of Surgery · Contact Numbers · Old Date (Old Date only for moved operations).
2. **Open the notifier in Google Chrome.**
   - If Claude in Chrome is available (`mcp__claude-in-chrome__*` tools, or a single
     `enable__mcp__claude-in-chrome` tool to call first), read the `chrome-browser` skill, call
     `tabs_context_mcp`, open a new tab with `tabs_create_mcp`, and `navigate` it to
     `https://claude.ai/artifact/PN4YWSqYjDKnqUJcgBekuz`. Leave the tab open for the user; do not
     close it at the end. Claude only opens the page: dropping the Excel and pressing Send stay
     with the user, and Claude never sends a WhatsApp message itself.
   - If Claude in Chrome is not connected or its calls fail, do not retry more than once. Give
     the user the link and tell them to open it in Google Chrome (not the Claude app or another
     browser), where they are signed in to WhatsApp Web.
   - If the artifact itself is gone (deleted, or owned by another account), publish
     `assets/ctsppum-notifier.html` as a new artifact with `icon: "message"`, then open that new
     link in Chrome as above. Load the `artifact-design` skill before any publish.
3. **Tell the user what to do on the page**, briefly:
   - Drop `patient_list.xlsx` on step 1.
   - Rows with an Old Date get the 🔄 change-of-date script automatically; others get the
     standard admission script.
   - Pick the global language at the top; any row can be pinned to its own language.
   - Edit the template or any single message before pressing Send / Alt.
   - Send / Alt open WhatsApp Web in a new Chrome tab with the message filled in; they check it
     and press send there. If WhatsApp Web asks for a QR code, they link Chrome once from
     WhatsApp on their phone (Linked devices).
4. Never paste the patient list into the chat or into the artifact source. The Excel stays on
   the user's device.

## Updating the page

- Edit `assets/ctsppum-notifier.html`, then publish it to the same URL (pass `url` from
  another conversation; read the artifact first, as the Artifact tool requires).
- Keep the date logic: `{{Date-2days}}` = ward admission, `{{Date-1days}}` = ward admission
  if the patient attends the Friday PAC clinic, `{{Date-5days}}` = stop antiplatelets / SGLT2i /
  anticoagulants, all computed from the NEW operation date.
- Phone numbers are normalised to `60…` (Malaysia). Multiple numbers split on `/ , ;`.
- The artifact frame blocks `window.open` for most viewers, so send buttons must stay real
  `<a href target="_blank">` links; refresh the href in the click handler so hand edits are sent.
- Keep the "Try with sample patients" data fictional.
- Bump the version label in the header and in `plugin.json` together.
