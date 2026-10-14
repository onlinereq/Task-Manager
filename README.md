# Task Manager

A vibe-coded web app that uses Google Sheets as its backend. Team members add tasks to a shared dashboard and update statuses so everyone can follow progress.

This app was built for small team use, as a lighter, more visual alternative to a traditional task board or spreadsheet. If you've found this repo, feel free to use it at your own risk.

## Features

- **Shared dashboard:** Everyone can add tasks and update their status.
- **Personal login:** Each user signs in with a profile name and PIN.
- **Google Sheets sync:** All data is saved to your own sheet through a web app URL.
- **Clear visuals:** See task progress at a glance, with easy navigation.

## Security

This app has basic login protection to keep strangers from viewing or editing the data. It is **not** built to high-end security standards, so don't use it for sensitive or confidential information.

## Setup

1. **Create or open your Google Sheet.**
   For a better view while setting things up, link Apps Script directly to a real sheet: open your sheet → **Extensions → Apps Script**, then paste in the `Code.gs` code.

2. **Run the setup function.**
   In the Apps Script editor, select the `setupSheet` function from the dropdown and click **Run**. This creates the required sheets and headers automatically.

3. **Add your team.**
   Open the generated Users sheet and add each team member's name and PIN.

4. **Save and deploy.**
   Save the script, then go to Deploy → New deployment → Web app. Copy the resulting URL.

5. **Connect the frontend.**
   Paste the URL into `config.js` (replace the placeholder URL).

6. Open `index.html` and start adding tasks.

> Using your own web app URL keeps all progress stored safely in your own Google Sheet.
