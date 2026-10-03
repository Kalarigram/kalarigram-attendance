# Kalarigram Attendance

Single-page check-in app (`index.html`, served by GitHub Pages from `main`) backed by a
Google Apps Script web app (`apps-script/Code.gs`, deployed separately in Apps Script).

## When finishing any change

Always end with clickable links so the owner can check the change:

- **Preview of the pushed branch** (works before merging):
  `https://raw.githack.com/Kalarigram/kalarigram-attendance/<branch>/index.html`
- **Live site** (only updates after merging to `main`):
  `https://kalarigram.github.io/kalarigram-attendance/`

Note that the preview talks to the real Apps Script backend, so check-ins and registrations
made there are saved to the real sheet. If `Code.gs` changed, say that it must be pasted into
the Apps Script project and redeployed (Deploy → Manage deployments → Edit → New version).
