# Like Button Web Part

## Summary

A SharePoint Framework (SPFx 1.22) web part that adds a configurable **Like** button to a modern SharePoint page. Users click it to like (and optionally unlike), and the like count updates immediately.

- **Icons**: thumbs up, heart, star, or any emoji (e.g. 👏 🎉 🚀)
- **Text**: a display text (e.g. *"Did you find this page helpful?"*) and an optional *"after liking"* text
- **Sizing**: Small, Medium, Large, Extra large
- **Layout**: text beside or above the button; left, centre or right aligned
- **Style**: Filled, Outline or Subtle, with an optional custom accent colour (otherwise follows the site / section theme)
- **Count**: show or hide, custom singular / plural labels (`1 like`, `12 likes`, `1.2K likes`), "Liked by …" tooltip on hover
- **Unlike**: allow or prevent users removing their like
- Optimistic update with rollback and an error message if saving fails; keyboard and screen-reader accessible (`aria-pressed`)

## Where likes are stored

Choose in the property pane, on the **Data** page:

| Storage | How it works | Permissions |
| --- | --- | --- |
| **Page's built-in likes** (default) | Uses the page's own likes via `/_api/sitepages/pages({id})/like`, so the count matches the page's social bar and "Like" in the page footer. One count per page. | Anyone who can read the page. |
| **SharePoint list** | One list item per user per *like key* (Title = key, Created By = user). Lets you have several independent Like buttons on a page, or share one count across pages by giving them the same key. Works in the workbench. | Users need **Contribute** on the list. The list (default `PageLikes`) is created automatically the first time someone with *Manage Lists* likes; or create a custom list with that name yourself. |

The default like key is the page URL plus the web part instance ID. The list is queried with a `Title eq` filter; if it may grow past 5,000 items, index the **Title** column.

## Compatibility

| Item | Version |
| --- | --- |
| SPFx | 1.22.2 (Heft build) |
| Node.js | 22.14+ (v22 LTS) |
| Hosts | SharePoint Online pages, full-width sections, Teams tab |

## Build and deploy

```bash
npm install
```

```bash
npm start
```

`npm start` serves to the hosted workbench (`https://{tenantDomain}/_layouts/workbench.aspx`; set the `SPFX_SERVE_TENANT_DOMAIN` environment variable or edit `config/serve.json`). Page likes need a saved site page, so switch Storage to **SharePoint list** to try it in the workbench.

```bash
npm run ship
```

This produces `sharepoint/solution/new-like-web-part.sppkg`. Upload it to the App Catalog, add **Like button** to a page, and open the property pane to configure it.

## Project layout

```
src/webparts/like/
  LikeWebPart.ts              web part + property pane
  models.ts                   property types
  components/LikeButton.tsx   UI, optimistic update, accessibility
  components/LikeIcon.tsx     SVG icons / emoji
  components/LikeButton.module.scss
  services/PageLikeService.ts page's built-in likes
  services/ListLikeService.ts SharePoint list storage
  loc/                        strings
```
