# Spotsylvania Yard Waste Alert

A demo of the Chancellor Convenience Center yard waste workflow. An attendant taps a status in an Apps Script web app, a private Google Sheet records the event, and a scheduled Apps Script function updates the status shown by the mock county website on GitHub Pages.

## Run the demo

1. Make a private Google Sheet for this demo. In the Sheet, open **Extensions → Apps Script**. Copy in `apps-script/Code.gs`, add an HTML file named `Staff` using `apps-script/Staff.html`, and enable **Show appsscript.json manifest file in editor** in Project Settings. Replace the manifest with `apps-script/appsscript.json` to limit spreadsheet access to this document.
2. Under **Project Settings → Script Properties**, set `OFFICE_KEY` to a unique random key of at least 32 characters. Keep the key with the demo operator. Run `setupDemo` from the script editor and approve the requested access. It prepares the log and published status tabs and installs the one-minute trigger.
3. Deploy as a **Web app**, with **Execute as** set to the deployment owner and access set to **Anyone** so the mock public page can be read signed out. The staff functions check the office key before reading or writing. Copy the deployed `/exec` URL.
4. Put the public `/exec` URL in `docs/config.js`. In the GitHub repository’s **Settings → Pages**, select **Deploy from a branch**, your default branch and **/docs**. The page needs the deployment URL only.
5. Open the `/exec` URL on the staff device and enter the office key to set up the browser. Tap **Full**. Verify the new row in **Status Log**, then the corresponding value in **Published Status** and the new alert on the Pages site in a signed-out browser. Repeat with **Accepting** to verify reopening. The staff app confirms when the update is queued. The scheduled publisher writes queued updates to the Sheet, then makes the latest status available to the public reader.

The demo displays **Current status unavailable** when it has no published status or the last confirmation is older than two hours. For a pilot, the county should choose the expiry window and apply its opening, seasonal and holiday schedule before showing a current acceptable status.

## County handoff

Fork or copy the source and create a new private Sheet in the county’s Google Workspace. Deploy the script there, choose staff access and install the trigger under a county-controlled account. In `publishPending_`, the demo currently writes the public projection to the **Published Status** tab. The county can use that transition to update its existing website alert through a connection its web team approves. Keep the release separate from the mock. The mock uses its own alert reader and does not contact the county site.

The source includes neither a working office key nor county credentials. Keep operator keys in Apps Script Script Properties and keep deployment credentials out of the repository.
