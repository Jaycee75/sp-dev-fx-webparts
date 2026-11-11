# Page Rollup

A SharePoint Framework web part that rolls up pages (or documents) from a site and library you choose and shows them as modern thumbnail tiles.

## What it does

- **Source**: any site in the tenant (blank = the current site) and any document library. Site Pages is picked automatically when the web part is added.
- **Column filters**: up to 3 filters on library columns (text, choice, managed metadata / tags, people, lookup, yes/no, number). Conditions are *is*, *is not*, *contains*, *is empty* and *has a value*. Separate several values with `;`. You can require all filters or any filter to match.
  - Example: library = Site Pages, filter = `Tags` *is* `Lesson Learning`.
- **Date range** on any date column (defaults to *First Published Date* for pages): any time, the last X days, this quarter, last quarter, or this year.
- **Sort** by newest, oldest, title or most viewed, with a **maximum number of items** from 1 to 50.
- **Pages-only options**: show only published pages, and only news posts.
- **Text on tiles** can be turned on or off. Pick up to 5 columns; the first one is the heading. You can also show the column names.
- **Page views** badge in the corner (all time or last 14 days), which can be turned on or off.
- **Six layouts**: Tiles, Cards, Hero, Mosaic, Filmstrip (scrolling with arrows) and List.
- **Three text styles**: dark gradient, frosted glass, or a solid theme-colour band.
- **Hover effects**: zoom the picture, lift the tile, reveal extra details, or do nothing.
- **Sizing**:
  - Height grows to fit, or is fixed with the tiles scaling to fill it.
  - Tiles per row is automatic (based on width) or set to 1–6.
  - Also adjustable: target tile width, picture shape (16:9, 4:3, 1:1, 3:4), spacing and corner roundness.
- Follows the site theme and section backgrounds, works on phones, and respects "reduce motion".

## Where the data comes from

| What | API | Permissions |
| --- | --- | --- |
| Items, columns, libraries | SharePoint REST (`RenderListDataAsStream`, `lists`, `fields`) | The viewer's own SharePoint permissions |
| Page views | SharePoint Search (`ViewsLifeTime`, `ViewsRecent`) | The viewer's own SharePoint permissions |

There are **no Microsoft Graph permission requests**, so the package can be deployed to a site collection app catalog without a SharePoint admin.

Notes:

- Column filters run in the browser over the newest items in the date range. Up to 2,000 items are checked, or 1,000 when sorting by most viewed.
- View counts come from the search index, so new pages may show 0 views for a while.
- Sorting by the date column on a library with more than 5,000 items needs that column to be indexed.

## Build

Requires Node.js 22 (`>=22.14.0 <23`).

```bash
npm install
```

```bash
npm run build
```

The package is created at `sharepoint/solution/page-rollup.sppkg`. Upload it to your site's app catalog, then add **Page Rollup** to a page. It's under the *Advanced* group in the web part picker.

## Version history

| Version | Date | Comments |
| --- | --- | --- |
| 1.0.0 | November 2026 | Initial release |
