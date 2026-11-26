# The Ephemera Cabinet: setup

Already done in your Google Drive, in the folder **Design Archive**:

- **Design Archive — Database** (Google Sheet) with these tabs:
  - **Works**: one row per designed work.
  - **Entities**: clients, each with a short code in `abbr`, e.g. BTR.
  - **Designers**: Individual, Print House or Studio, with years of activity.
  - **Solutions**: one row per type (columns: category, category_code, subcategory, subcategory_code, type). **Add a row here to add a new option.** It appears in the Works dropdowns straight away.
  - **Styles**: one column per movement, with the movement name in row 1 and its styles below. Add styles anywhere in a column; to add a movement, fill in the next empty header.
  - **_lists** (hidden): works out each row's allowed options. Don't edit it.
- **Originals (full quality)** and **Web (compressed)** folders.
- 24 sample records, marked "Sample record".

## Dropdowns in Works

| Column | Lists | Comes from |
|---|---|---|
| category | all categories | Solutions |
| subcategory | only the subcategories of that row's category | Solutions |
| types | every type, as one rule for the whole column with multi-select; the script removes types that don't belong to the row's subcategory | Solutions, column E |
| movement | all movements | Styles, row 1 |
| design_style | only the styles under that row's movement | Styles |
| typography_style | every style, from all movements | Styles |

- **Turn on multi-select for types (one time).** Select J2:J1025, choose Data → Data validation, click the rule, open Advanced options, tick **Allow multiple selections**, then click Done. Column J has a single rule for the whole column, so this setting stays put.
- **Types that don't fit.** If you pick a type that doesn't belong to the row's subcategory, the script removes it and shows a short message.
- **Leave columns I and N alone in the Data validation panel.** Subcategory and design_style have one rule per row so that their lists can depend on the category and movement. Editing them there merges the rules into one. If that happens, run **Archive → Repair / extend dropdowns**. That repair also runs by itself when rows are inserted or deleted, and it never touches column J.
- **Row coverage.** Rows 2–51 have the per-row lists now. The script keeps 100 rows ready past your last entry.

## Work IDs

```
PE  - OPA - HBV - 0004 - 1938
cat   sub   client uid   year
```

- The IDs are built by the script; don't type them.
- `uid` is a permanent number given to each new row.
- Change a row's category, subcategory, client or year and the ID is rebuilt, while the uid and the linked images stay the same.
- The category and subcategory codes come from the Solutions tab (`category_code` and `subcategory_code`). If a code is left blank, a three-letter code is made from the name.
- Client codes come from Entities → `abbr`. Changing one rebuilds every ID that uses it.

## 1. Install the script (once)

1. Open the sheet and choose **Extensions → Apps Script**.
2. Paste in `Code.gs`. The folder IDs are already filled in.
3. In the left panel, click **Services +** and add **Drive API**.
4. Save and reload the sheet. An **Archive** menu appears.
5. Choose **Archive → Repair / extend dropdowns**, then **Archive → Rebuild all work IDs**, then **Archive → Turn on automatic sync + dropdown repair**. Google asks for permission once.

## 2. Publish the data feed

1. In Apps Script, choose **Deploy → New deployment → Web app**.
2. Set **Execute as: Me** and **Who has access: Anyone**.
3. Copy the URL that ends in `/exec`.
4. Paste that URL into `const DATA_URL = "";` in `index.html`.

## 3. Add images

There are two ways to connect a scan to its record:

- Name the scan starting with its work ID, e.g. `PE-OPA-HBV-0004-1938 menu front.jpg`, and drop it into Originals.
- Or drop in any file name. The next sync creates a draft row for it, with the file name already in `image_file`. Fill in the record, then tick **publish**.

For each image, the sync then:

- shares the original as view-only, by link
- saves a 1600 px JPEG copy into Web, named by uid (`0004_web.jpg`), so it survives ID changes
- writes both links, the original's pixel size and the file sizes into the row

## 4. Host the site

The claude.ai version can only show the sample records. To go live, upload `index.html` to GitHub Pages or Netlify Drop.

## On the site

- Each **drawer** is a category.
- Each **black guide** is a subcategory. Hover it (or tap it on a phone) to fan out papers for up to 7 of its works. Click a paper to open that work.
- Hovering a **coloured tab** lifts that single folder.
- Each **entry** shows the image on the left. On the right are the lined-paper record and sticky notes for designer, client, design philosophy and solution.
- Each **type** (Menus, Napkins…), design style and typography style has its own page.
