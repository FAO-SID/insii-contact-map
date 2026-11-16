# INSII focal contact map

Interactive map for members of the International Network of Soil Information Institutions (INSII) to check the registered INSII institution and contact name for each country. The basemap is UN ClearMap (official UN boundaries). The page is one static file, `index.html`.

## What it shows

- Country dropdown (grouped by region). Selecting a country zooms to it and fills two cards: INSII institution and INSII contact name.
- One dot per country at the capital. Clicking a dot opens a popup with the same two fields.
- Direct links per country: `index.html#ERI` opens Eritrea (ISO3 code).
- Only institution and contact name are shown. Email addresses are never read or published.

## Repository contents

| Path | Purpose |
|---|---|
| `index.html` | Built page (map library, data and code in one file). This is what gets published. |
| `build.py` | Rebuilds `index.html` from the Excel workbook. |
| `src/template.html` | Page source (layout, styling, map logic, disclaimer text). |
| `data/country_bbox.json` | Country extents, used only to set the zoom. Not drawn on the map. |
| `vendor/maplibre-gl-4.7.1/` | MapLibre GL JS (BSD-3-Clause), inlined into the build. |
| `docs/original_powerbi_draft.png` | Screenshot of the earlier Power BI draft, for reference. |
| `THIRD_PARTY_NOTICES.md` | Licences and data credits. |

The Excel workbook is not part of the repository and must not be committed. `.gitignore` excludes `*.xlsx`.

## Updating the contacts

1. Update the `Website` sheet in the INSII workbook and save it in Excel (the build reads the values Excel last saved).
2. Run:
   ```
   pip install openpyxl
   python3 build.py path/to/workbook.xlsx
   ```
3. The script prints the number of countries and lists any gaps (missing institution, missing zoom extent). It stops if an email-like string appears in the data or if an ISO code is duplicated.
4. Commit and push the new `index.html`.

Required columns in the `Website` sheet (matched by header name): first column ISO3 code (header may be blank), `Country`, `Lat`, `Lon`, `Region`, `Sub Region`, `INSII Contact`, `INSII (Institution)`. Values "NA", "N/A" and "Unspecified" are shown as "Not specified". Lat/Lon are the capital-city coordinates where the dot is placed.

## Publishing on GitHub Pages

1. Create a repository (preferably under an FAO or OCB organisation account) and push the contents of this folder to the `main` branch:
   ```
   git init
   git add .
   git commit -m "INSII focal contact map"
   git branch -M main
   git remote add origin https://github.com/<account>/<repository>.git
   git push -u origin main
   ```
2. In the repository: Settings > Pages > Build and deployment > Source "Deploy from a branch" > Branch `main`, folder `/ (root)`.
3. The site appears at `https://<account>.github.io/<repository>/` within a few minutes.

GitHub Pages on a private repository needs a paid GitHub plan. On a free account the repository, and therefore the contact data inside `index.html`, is public.

## Basemap: UN ClearMap

- Style and tiles are loaded live from the UN server when the page opens:
  `https://pro-ags2.dfs.un.org/arcgis/rest/services/Hosted/Clear_Map/VectorTileServer`
  (style: `.../resources/styles/root.json`). The page therefore needs internet access.
- Boundaries, labels and relief are drawn exactly as served. No layer is edited. Only the relative links in the style were replaced with full addresses of the same service.
- The style has no layers above zoom 5.9. The page caps zoom at 5.85, so very small countries do not zoom in tightly.
- The service was open without login when tested in November 2026 and allows cross-origin requests.
- The boundary tiles carry no country-code field, so the selected country cannot be shaded. The selected dot and the zoom mark it.

## Open items before publishing publicly

- Clearance from FAO data protection or communications to publish contact names on a public page.
- Confirmation from UN Geospatial or FAO GIS that a public FAO page may use ClearMap. The service page states no terms of use.
- Confirmation of the map disclaimer wording (included in `src/template.html`, left panel).
- Test of the live page from the final `https` address in the browsers members use.

## Notes on the data

- Source: `Website` sheet of the GSP/INSII contact network workbook. 113 countries at the time of the first build.
- Four countries have no institution recorded (Azerbaijan, Chile, Ethiopia, Finland) and display "Not specified".
- The `EU` row has no country shape and the map flies to a fixed view.
- The page includes `noindex` so search engines are asked not to list it. This does not restrict access.
