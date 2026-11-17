# Bump & Binge: setup

A shared movie watchlist for a coworker's maternity leave. It's two pieces:

- **`index.html`**: the page people visit, hosted free on GitHub Pages.
- **`Code.gs`**: a small Google Apps Script that saves the movies into a Google Sheet you own.

Nobody needs an account to add a movie. About 15 minutes, start to finish.

---

## 1. Make the Google Sheet and its script

1. Go to [sheets.new](https://sheets.new) and name the sheet something like **Bump & Binge**.
2. In the menu, open **Extensions → Apps Script**.
3. Delete what's in the editor, paste in all of `Code.gs`, and click **Save** (the disk icon).
4. Click **Deploy → New deployment**.
   - Next to "Select type", click the gear and choose **Web app**.
   - **Execute as:** Me
   - **Who has access:** Anyone
   - Click **Deploy**.
5. Google will ask you to authorize. Choose your account. If you see "Google hasn't verified this app", click **Advanced → Go to (project name)**. This is normal for your own scripts. Then click **Allow**.
6. Copy the **Web app URL**. It looks like `https://script.google.com/macros/s/AKfy…/exec`.

## 2. Connect the page to the sheet

In this repository on GitHub, open `index.html`, click the ✏️ pencil to edit, and find this line near the bottom:

```js
const SCRIPT_URL = "";
```

Paste your URL between the quotes:

```js
const SCRIPT_URL = "https://script.google.com/macros/s/AKfy…/exec";
```

Click **Commit changes**. Or send the URL to Claude and it'll make the change.

## 3. Turn on the link with GitHub Pages

1. In this repository, go to **Settings → Pages**.
2. Under "Build and deployment", set **Source: Deploy from a branch**, pick the branch that has `index.html` in it, set the folder to **/ (root)**, and click **Save**.
3. After a minute or two, the link appears at the top of that page:
   **`https://bbmjps94ww-ux.github.io/bump-and-binge/`**

Share that link. Done!

---

## Good to know

- **Adding a test movie:** add one yourself from the live page, check it shows up in the sheet's **Recs** tab, then delete that row.
- **Removing or fixing a movie:** edit or delete the row in the sheet. The page shows the change on its next refresh.
- **Updating the script later:** after editing `Code.gs` in Apps Script, use **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. That keeps the same URL. A *New deployment* would make a new URL.
- **Privacy:** anyone with the link can read the list and add to it. The sheet itself stays private to you unless you share it.
- **Want your coworker to see the raw list?** Share the Google Sheet with her as a viewer, or just send her the page link.
