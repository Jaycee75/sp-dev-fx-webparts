# Page Rollup

A SharePoint Framework web part that rolls up pages (or documents) from a site and library you choose and shows them as modern thumbnail tiles.

## What it does

- **Heading**: type it straight onto the page in edit mode, like SharePoint's own web part titles. A small toolbar above it sets the size (S, M, L, XL) and alignment (left, centre, right). The heading is hidden on the published page when it's empty.
- **Source**: any site in the tenant (blank = the current site) and any document library. Site Pages is picked automatically when the web part is added.
- **Column filters**: up to 3 filters on library columns (text, choice, managed metadata / tags, people, lookup, yes/no, number). Conditions are *is*, *is not*, *contains*, *is empty* and *has a value*. Separate several values with `;`. You can require all filters or any filter to match.
  - Example: library = Site Pages, filter = `Tags` *is* `Lesson Learning`.
- **Date range** on any date column (defaults to *First Published Date* for pages): any time, the last X days, this quarter, last quarter, or this year.
- **Sort** by newest, oldest, title or most viewed, with a **maximum number of items** from 1 to 50.
- **Pages-only options**: show only published pages, and only news posts.
- **Text on tiles** can be turned on or off. Pick up to 5 columns; the first one is the heading. You can also show the column names.
- **Views, likes and comments**: each can be turned on or off. They're shown together in a badge in the corner of each tile, under the picture on the Card carousel, and as a line of text in the List layout.
  - Views come from SharePoint search (all time or last 14 days).
  - Likes and comments come from the page's own like/comment counts.
  - For document libraries, turn on *Library settings > Rating settings > Likes* to give them a likes count.
  - You can also sort by *Most liked*.
- **Nine layouts**:
  - **Tiles**: equal tiles with text over the picture.
  - **Cards**: picture on top, text underneath.
  - **Hero**: the newest item as one big feature tile, the rest smaller beside and below it.
  - **Mosaic**: a magazine-style mix of wide and tall tiles, with no single feature tile.
  - **Filmstrip**: one row you scroll sideways.
  - **List**: compact rows with a small thumbnail.
  - **Spotlight carousel**: the current slide in the middle with its neighbours peeking at the sides. The background can be a blurred copy of the current picture, a picture you choose (blurred or sharp), a gradient, a solid colour, a light theme colour, or none. Picture backgrounds can be darkened so the slides stand out.
  - **Card carousel**: social-media style post cards (header, picture, views, caption) sliding along. Column 2 is the card header (e.g. author) and column 3 the line under it (e.g. date).
  - **Slide-up ticker**: a column of tiles; the top one slides up and away and the next rises in from the bottom.
- **Carousel motion**: move on automatically (seconds per item), slide speed, pause on hover, loop, arrows and dots. Every carousel can also be swiped on touch screens, dragged with the mouse, clicked, or moved with the arrow keys.
- **Picture quality**: *Sharp* (default) uses SharePoint-resized copies sized for each tile, at high-DPI resolution, and switches to the full-size picture if a copy comes back too small. *Faster loading* uses smaller copies, and *Original* always uses the full-size picture. Banner images that SharePoint stored as small previews are unwrapped to the original picture.
- **Three text styles**: dark gradient, frosted glass, or a solid theme-colour band.
- **Hover effects**: zoom the picture, lift the tile, reveal details, or do nothing. With *Reveal details* you pick a **details column** (e.g. Description). Its full text, line breaks included, appears across the middle of the tile, card or slide on hover, and long text scrolls. Works in every layout, including the carousels.
- **Sizing**:
  - Height grows to fit, or is fixed with the tiles scaling to fill it.
  - Tiles per row is automatic (based on width) or set to 1–6.
  - Also adjustable: target tile width, picture shape (16:9, 4:3, 1:1, 3:4), spacing and corner roundness.
- Follows the site theme and section backgrounds, works on phones, and respects "reduce motion".

## Document libraries

Point the web part at any document library, or add **Document Rollup** from the toolbox. It's the same web part, but starts on the *Documents* library with cards, a 4:3 picture shape and *Modified* as the date column.

- **Previews**: Word, Excel, PowerPoint, PDF, Visio, text files, videos and pictures get a preview of the file. The web part tries these sources in order:
  1. the thumbnail service SharePoint uses for its own library tiles view
  2. the preview handler, addressed by site/web/file IDs
  3. the preview handler, addressed by path
- **No preview available**: the tile shows an icon in the file type's colour (Word blue, Excel green, PDF red, and so on).
- **Picture fit**:
  - *Fill the tile* (the default) crops from the top of the first page, so the document's heading shows.
  - *Show the whole picture or page* fits the entire page in the tile.
- **Opening files**: Office files open in Office for the web, and other files open directly.
- **Titles**: documents with no Title use their file name.

## Where the data comes from

| What | API | Permissions |
| --- | --- | --- |
| Items, columns, libraries | SharePoint REST (`RenderListDataAsStream`, `lists`, `fields`) | The viewer's own SharePoint permissions |
| Page views | SharePoint Search (`ViewsLifeTime`, `ViewsRecent`) | The viewer's own SharePoint permissions |
| Likes, comments | Library columns `_LikeCount`, `_CommentCount` (pages) or `LikesCount` (libraries with ratings set to Likes) | The viewer's own SharePoint permissions |

There are **no Microsoft Graph permission requests**, so the package can be deployed to a site collection app catalog without a SharePoint admin.

Notes:

- Column filters run in the browser over the newest items in the date range. Up to 2,000 items are checked, or 1,000 when sorting by most viewed.
- View counts come from the search index, so new pages may show 0 views for a while.
- Sorting by the date column on a library with more than 5,000 items needs that column to be indexed.

## Performance

- **Fast start-up:** the web part draws straight away and loads data in the background. It doesn't hold up page loading or SharePoint's web part start-up timing.
- **No phantom edits:** default settings are applied when read, not written back. Opening a page in edit mode doesn't mark the web part as changed, so it doesn't create a new draft.
- **Shared caches:** library lists, columns and site IDs are shared between every Page Rollup on the page and kept for 15 minutes in the browser session. Results are kept for 5 minutes and shown instantly, then refreshed if more than a minute old.
- **Fewer requests:** view counts load after the tiles are shown. Without filters, only the rows needed are read. With filters, rows are read in batches of 500.
- **Lighter pictures:**
  - *Sharp*, the default, uses SharePoint-resized copies sized for each tile, and switches to the full-size picture only if a copy comes back too small.
  - Carousels only load pictures for slides near the screen.
  - Autoplay pauses while the carousel is scrolled out of view.

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
| 1.1.0 | November 2026 | Spotlight, card and slide-up carousels; sharper images; new mosaic pattern |
| 1.6.0 | November 2026 | Performance: non-blocking start-up, no default writes (no phantom drafts), shared caches, deferred view counts, sized pictures, lazy carousel images, off-screen autoplay pause |
| 1.5.0 | November 2026 | Likes and comments counts (with Most liked sort) alongside views |
| 1.4.0 | November 2026 | Inline-editable heading with size and alignment toolbar |
| 1.3.0 | November 2026 | Document previews (thumbnail API with fallbacks), file-type placeholders, picture fit option, Document Rollup toolbox entry |
| 1.2.0 | November 2026 | Hover details column (full text) in all layouts; picture background for Spotlight; card carousel arrows moved beside the dots |
