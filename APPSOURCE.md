# Publishing Takeoff to Microsoft AppSource

Requirements change, so check Microsoft's current "Office Add-ins" publishing documentation in
Partner Center before you submit. This is the practical path.

## 1. Build
```
npm install                 # adds esbuild so app.js is minified
npm run vendor              # downloads pdf.js + dxf-parser into public/vendor (needs internet)
npm run build -- https://addin.yourdomain.com --email support@yourdomain.com --provider "Your Company"
npm run validate            # checks dist/manifest.xml with Microsoft's validator
```
Upload everything in `dist/public` to your HTTPS host (Azure Static Web Apps, Netlify, GitHub Pages, etc.).
`dist/manifest.xml` is what you sideload to test and what you submit.

The manifest adds a **Takeoff** button on the Excel **Home** tab. Before submitting, also change the
`<Id>` GUID if you want a new identity, and decide the final display name and description.

## 2. Test the hosted build (all three, not just desktop)
- Excel on Windows, Excel on the web, Excel on Mac: sideload `dist/manifest.xml`.
- The Home-tab button opens the pane. Open a DXF and a PDF (vector and scanned).
- Scale by preset, by calibration, per page, per region. Send a row and a table to cells.
- Save, close and reopen the workbook: scales and takeoff items return.
- Block the CDN (offline): everything still loads, because the libraries are in `app.js`.

## 3. Listing checklist (Partner Center)
- Listing logo: `store-assets/icon-300.png` (300x300).
- Screenshots: 1 to 5 real screenshots of the pane in Excel (not provided; take them from a real run).
- Name, short summary, long description, keywords, category.
- Privacy policy URL: `https://your-host/privacy.html` (a plain-language template. Have it reviewed
  for your situation: it is not legal advice).
- Support URL: `https://your-host/support.html`. Terms of use: optional but recommended.
- Pricing: a free listing is simplest. Paid add-ins need a licensing or SaaS offer setup.
- Notes for the reviewer: no sign-in needed, plus 2 or 3 sample drawings (a small DXF and a vector PDF)
  and the steps above, so they can see it work.

## 4. Common reasons for rejection
- Any `localhost` or non-HTTPS URL left in the manifest or hosted files.
- Missing or broken support and privacy pages.
- Ribbon button missing, or the pane failing on one platform (test web and Mac).
- Description or name that implies Microsoft endorsement.

## Before each release
- Bump `<Version>` in manifest.xml (and "version" in package.json).
- Re-test: Home-tab Takeoff and Takeoff help buttons, every measuring shape, send to a selected cell,
  a fixed cell, a linked cell, a table on a new sheet, and "Insert view as picture".
