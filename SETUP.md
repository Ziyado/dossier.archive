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

## 3. Add images (any number per work)

Images live on their own **Images** tab, one row per image:

| Column | What it holds |
|---|---|
| uid | The work's permanent number. The script fills it in. |
| work_id | The work the image belongs to, picked from a dropdown. It updates automatically if the work's ID changes. |
| order | 1 is the main image; 2, 3 and so on follow on the entry page. |
| image_file | The file's exact name in Originals. |
| caption | Optional, e.g. Front, Back, Inside, Detail. |
| grey columns | Filled by the script: the Drive IDs, the full and web links, the size in pixels and a status line. |

Ways to add images to a work:
- **By name:** start the file name with the work ID and drop it into Originals, e.g. `PE-OPA-HBV-0004-1938 front.png` and `PE-OPA-HBV-0004-1938 back.png`. Each file becomes its own row on the Images tab, numbered 1, 2, 3 in the order they're found. Edit `order` afterwards if you want a different sequence.
- **By hand:** add a row on Images, pick the `work_id`, type the file name into `image_file`, and set `order`.
- **Image first:** a file with any other name gets a new draft work on Works, plus its own row on Images. Fill in the work, and add more images to it with either of the methods above.

Each image gets its own web copy, named by uid and position (e.g. `0004-02_web.png`), in the original's format. The `status` column on Works shows how many images each work has.

On the site, the entry page shows image 1 first. Arrows and a row of thumbnails switch between images, and **Full resolution** always opens the image currently shown. In the drawers, papers for works with more than one image carry a small "3 images" label.

To remake web copies, select rows on Images (or works on Works, which covers all their images) and run **Archive → Remake web copies for selected rows**.

## 4. Host the site

The claude.ai version can only show the sample records. To go live, upload `index.html` to GitHub Pages or Netlify Drop.

## On the site

- Each **drawer** is a category.
- Each **black guide** is a subcategory. Hover it (or tap it on a phone) to fan out papers for up to 7 of its works. Click a paper to open that work.
- Hovering a **coloured tab** lifts that single folder.
- Each **entry** shows the image on the left. On the right are the lined-paper record and sticky notes for designer, client, design philosophy and solution.
- Each **type** (Menus, Napkins…), design style and typography style has its own page.
