# Quote Generator — One-Click Windows App

One app handles **all quote types** (Personal Auto and Homeowners HO3 so far). It detects which kind of export was loaded from the spreadsheet's columns and uses the matching proposal template automatically — staff never choose anything.

## ⚠️ Before uploading to GitHub

Two template files are NOT in this folder yet (they couldn't be copied automatically). Drop them in from your computer:

1. `Proposal  PERSONAL AUTO 23mar2026.pptx` (note: two spaces after "Proposal")
2. `HO3 Proposal 01july2026.pptx`

The build will fail without them.

## Getting the installer (built free in the cloud — no Python needed)

1. Create a free account at **github.com**, then click **+ → New repository**. Name it `quote-generator`, set it to **Private**, click **Create repository**
2. Click **uploading an existing file** (link on the empty-repo page). Drag in EVERYTHING in this folder — including the `.github` folder (if drag-and-drop skips it, see note below). Click **Commit changes**
3. Open the **Actions** tab → click **Build Windows app** in the left list → **Run workflow** button → green **Run workflow**
4. Wait ~5 minutes for the green checkmark, click the finished run, scroll to **Artifacts**, download **QuoteGenerator-Setup**
5. Unzip it → `QuoteGenerator-Setup.exe` is your installer. Staff double-click it once, click Install, and get a **Quote Generator** icon on their Desktop and Start menu

**Note on the `.github` folder:** browsers sometimes don't upload hidden folders by drag-and-drop. If after upload you don't see `.github/workflows/build-windows.yml` in the repo, create it manually: **Add file → Create new file**, type `.github/workflows/build-windows.yml` as the name (the slashes create the folders), and paste the contents of that file from this folder.

**Updating later:** change any file (new template, mapping fix) → upload it to the repo again ("Add file → Upload files", it replaces the old one) → the build re-runs automatically → download the new installer from Actions.

## What staff see

1. Double-click the **Quote Generator** icon
2. The app finds the newest JotForm export by itself (Downloads/Desktop) and shows a plain list of **customer names** — nothing else
3. Click a name → optionally *add a cover photo* (drops a custom picture into the front page's existing frame — layout never changes) → press the big **Generate Quote** button
4. **"✓ Quote Created Successfully"** appears with an **Open File** button

Quotes are saved to the **`Quotes` folder on the Desktop**. Wrong file picked? The app says "Wrong file selected — please export from JotForm again." No technical errors ever appear; details go silently to `Desktop\Quotes\error_log.txt`.

## Setting up the HO3 (homeowners) quote type

1. **Create the JotForm form**: paste the prompt in `jotform_ho3_prompt.txt` into JotForm's AI builder. Field labels must stay exactly as written
2. **Test one quote**: do a test submission, export to Excel, generate. Check the coverage tables (page 4), premium summary (page 5), and options page (page 3). Any surviving sample value (e.g. "Holly White", "$ 9,156") means a mapping fix is needed — `inspect_template.py` dumps the template text for comparison

## Carrier logos (customizable)

The page-4 coverage tables show a carrier **logo image**. To make it match each quote's carrier:

1. Create a folder named **`CarrierLogos`** inside the **`Quotes` folder on the Desktop** (or next to `QuoteGenerator.exe`)
2. Save each carrier's logo there as a picture named after the carrier — `Chaucer.png`, `Lloyds.png`, `Citizens.jpg`, … (capitalization, spaces, and punctuation don't matter: `lloyd's` matches `Lloyds.png`)

When generating, the app swaps the logo to match the **Expiring Carrier** (top table) and **New Carrier** (Option A table). If no matching logo file exists, the old wrong logo is removed so the carrier *name* shows in the cell instead.

## Images straight from the JotForm (no clicks in the app)

The HO3 form can include three optional upload fields: **Cover Photo**, **New Carrier Logo**, and **Expiring Carrier Logo** (see Section 6 of `jotform_ho3_prompt.txt` — add them via JotForm's AI or manually, labels exact). JotForm exports each upload as a link; at generation time the app downloads and places them automatically. Priority order: photo chosen in the app → photo uploaded on the form → (for logos) the `CarrierLogos` folder → logo removed.

Two requirements for this to work: the PC needs internet when generating, and uploaded files must be viewable by link — in JotForm go to **Settings → Form Settings → Show More Options** and set **"Require log-in to view uploaded files" to No**. If a download fails for any reason, the quote still generates (that image is simply skipped) and the reason is logged to `error_log.txt`.

Other note: the Named Insureds box shows the single client name.

## Project files

| File | Purpose |
|---|---|
| `app.py` | The window — auto-detect, name list, cover photo, Generate Quote |
| `data_loader.py` | Reads/validates JotForm exports; detects quote type; finds newest export |
| `quote_profiles.py` | One profile per quote type: columns, template pattern, detection rules |
| `proposal_generator.py` | Fills templates — Auto path unchanged; HO3 via mapping JSON + table filling |
| `ho3_mapping.json` | HO3 template text → field map |
| `inspect_template.py` | Admin helper: dumps all text inside a .pptx template |
| `jotform_ho3_prompt.txt` | Paste into JotForm AI to create the matching HO3 form |
| `.github/workflows/build-windows.yml` | Cloud build recipe (exe + installer) |
| `installer.iss` | Recipe for the double-click installer |
| `build.bat` | Optional local build alternative (needs Python on your PC) |
