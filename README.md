# Takeoff: Excel add-in

Measure DXF and PDF drawings (vector or scanned) and send quantities into Excel as live formulas.

## Run it
1. `npm run certs`    one time: trusts the localhost HTTPS certificate
2. `npm run vendor`   optional for dev: downloads pdf.js and dxf-parser into public/vendor (offline use)
                      (without it, the libraries load from a CDN while developing)
3. `npm start`        serves the task pane at https://localhost:3000
4. `npm run sideload` opens Excel with the add-in loaded

Excel on the web: Insert > Add-ins > Upload My Add-in > `manifest.xml`.

## Release build (AppSource)
`npm install`, `npm run vendor`, then
`npm run build -- https://addin.yourdomain.com --email you@yourdomain.com --provider "Your Company"`
builds `dist/`: one bundled (and minified) `app.js` with the libraries inside, no CDN needed, plus a
manifest with a Home-tab ribbon button, icons, and support and privacy pages. See APPSOURCE.md for
the submission checklist.

## What it does
- **Files:** ASCII DXF (LINE, LWPOLYLINE/POLYLINE with curved segments, CIRCLE, ARC, ELLIPSE,
  SPLINE with fit points, INSERT block counts) and PDF (vector lines extracted; scans measured by hand).
- **Scales:** preset paper scale, two-point calibration, or a typed factor, applied to this page, a
  dragged region, chosen pages ("1-3, 5") or all pages. Priority: region > pages > all.
- **Length:** straight line, polyline, arc, Bézier curve, spline, freehand.
- **Area:** polygon, rectangle, rotated rectangle, circle, ellipse, sector, arc segment, freehand.
  Curved shapes use exact formulas. Area items can also output perimeter or volume (× depth) and can be deducted.
- **Count:** markers, or posts along a path at a spacing. **Angle.** **Symbols:** find look-alikes of a boxed symbol.
- **Pick:** click a vector line or closed shape to take its exact length or area from the drawing.
- **Editing:** click an item in Pan mode, drag its handles. Undo/redo, snapping, Shift to lock axis, live values while drawing.
- **Civil/BOQ:** code, repeat ×N, waste %, rate, width/depth for lengths (area, volume), notes.
- **Where quantities go:** selected cell or a fixed cell, move down/right/stay, quantity only or label+quantity+unit,
  chosen table columns (with Amount = Quantity × Rate and a total), a new Takeoff sheet, or items linked to exact
  cells in your own BOQ (Update linked cells). Insert the current view as a picture.
- **Check scale:** after a Distance measurement, type the known size to see how far off the scale is.
- **Saved with the workbook:** scales, items and links return when you reopen the same file.

## Limits
- Presets assume the PDF page is printed at the sheet size the scale was drawn for. If a sheet was
  shrunk (a 24x36 set saved as 11x17), use Calibrate.
- Symbol search finds same-size, same-orientation copies only. Check the boxes before adding the count.
- PDF layers are grouped by color. Filled shapes also report their outline as length, and a white
  page-background fill shows up as a big area (untick that layer).
- No DWG, no standalone image files, DXF block contents are counted but not measured.
- PDF extraction targets pdf.js 3.11.x (v4 changed the drawing-command format).
- Not yet run inside Excel by the author: test sideloading, session restore and table output on your setup.
