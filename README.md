# Meeting poll: setup

Two files do the work. `index.html` is the page people see, hosted on GitHub Pages. `Code.gs` is the backend, a Google Apps Script bound to one Google Sheet. One deployment serves every poll you create.

## 1. Try it first (no setup)

Open `index.html` in a browser. With `API_URL` empty the page runs in demo mode and stores everything in that browser. Create a poll, answer it, and look at the results before you deploy.

## 2. Backend (about 10 minutes, once)

1. Create a new Google Sheet, e.g. *Meeting polls*.
2. **Extensions > Apps Script**. Delete the placeholder code and paste in `Code.gs`. Save.
3. **Project Settings** (gear icon) > **Script properties** > **Add script property**:
   `CREATE_KEY` = a passphrase only you know. The page asks for it when you create a poll, which stops strangers using your Sheet.
4. Back in the editor, choose `setup` from the function menu and click **Run**. Approve the permissions. This creates the `Polls` and `Responses` tabs.
5. **Deploy > New deployment** > type **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Copy the web-app URL (it ends in `/exec`).

## 3. Front end

1. In `index.html`, paste the URL into `const API_URL = '';` near the top of the script.
2. Commit `index.html` to your GitHub Pages repository, replacing the old poll.
3. Open the site, enter your passphrase, and create a poll.

## Using it

- **Create:** open the site with no `?poll=` in the address. Pick dates, then start times; every date pairs with every start time, and you can remove any pair. List required attendees if some people must be there.
- **Share:** send the share link. Keep the admin link to yourself; it's also remembered on the device you used.
- **Respondents** mark each option Yes, If need be, or No. Their browser remembers them, so returning to the link lets them edit rather than add a second row.
- **Organiser tools** (admin link only): confirm the final time, edit required attendees, remove an entry, close the poll, download a CSV.

## How the ranking works

A slot drops out if any required attendee answers No. The remaining slots rank by Yes count, then by Yes + If need be, then by date. The chart shows every slot in order; faded bars are ruled out, and the outlined bar is the current best.

## Updating the backend later

After editing `Code.gs`, use **Deploy > Manage deployments > Edit > Version: New version**. This keeps the same URL. A *new deployment* would give a new URL, and you'd have to update `index.html`.

## Limits and known behaviour

- Up to 60 time options and 150 responses per poll. Fifteen or fewer options gets more careful answers.
- Time options can't be changed after creation; create a new poll instead.
- Someone answering from a second device can't reuse their name. The page tells them to ask you to remove the old entry or to add a surname.
- All times are SAST. If you ever poll people abroad, this needs time-zone handling added.
- Apps Script replies in 1–3 seconds. That's fine for this use but noticeably slower than a commercial service.
