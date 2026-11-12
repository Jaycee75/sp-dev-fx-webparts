import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IRollupField, IRollupItem, IRollupList, IRollupQuery } from '../models/IPageRollupProps';
import { getDateWindow } from '../utils/dateRange';
import { getFieldText, getFieldTextMultiline, rowMatches } from '../utils/fieldValues';

/* eslint-disable @typescript-eslint/no-explicit-any */

/** Site Pages libraries use this list template. */
export const SITE_PAGES_TEMPLATE = 119;

/** Columns that are hidden but still useful on Site Pages libraries. */
const USEFUL_HIDDEN_FIELDS = ['FirstPublishedDate', 'Description', 'BannerImageUrl', 'PromotedState', 'FileLeafRef', 'Modified', 'Created', '_LikeCount', '_CommentCount', 'LikesCount'];
/** Like and comment count columns: pages use _LikeCount/_CommentCount; libraries with ratings set to Likes use LikesCount. */
const LIKE_FIELDS = ['_LikeCount', 'LikesCount'];
const COMMENT_FIELDS = ['_CommentCount'];

const SUPPORTED_FIELD_TYPES = [
  'Text', 'Note', 'Choice', 'MultiChoice', 'Boolean', 'Number', 'Currency', 'Integer', 'Counter', 'DateTime',
  'User', 'UserMulti', 'Lookup', 'LookupMulti', 'TaxonomyFieldType', 'TaxonomyFieldTypeMulti', 'URL', 'Thumbnail', 'File'
];

const ALWAYS_SELECT = ['ID', 'FileRef', 'FileLeafRef', 'FileDirRef', 'File_x0020_Type', 'Title', 'UniqueId'];
const DEFAULT_PAGE_IMAGE = 'sitepagethumbnail.png';
const IMAGE_TYPES = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp'];
const PREVIEWABLE = ['docx', 'doc', 'docm', 'dotx', 'pptx', 'ppt', 'pptm', 'ppsx', 'xlsx', 'xls', 'xlsm', 'xlsb', 'pdf', 'rtf', 'odt', 'odp', 'ods', 'txt', 'csv', 'vsdx', 'one', 'mp4', 'mov', 'm4v', 'wmv', 'webm', 'heic', 'tif', 'tiff', 'svg'];

interface IWebIds {
  host: string;
  siteId: string;
  webId: string;
}
const OFFICE_TYPES = ['docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls', 'one'];

export interface IRollupResult {
  items: IRollupItem[];
  scanned: number;
  truncated: boolean;
  /** True when view counts still need loading with loadViews() (they are fetched after the tiles are shown). */
  viewsPending?: boolean;
}

// ---------------------------------------------------------------------------
// Caches shared by every Page Rollup on the page (and kept for a while in sessionStorage),
// so several web parts, or the same one re-created when switching edit/view, don't refetch.
// ---------------------------------------------------------------------------

const META_TTL = 15 * 60 * 1000; // libraries, columns, site ids
const RESULT_TTL = 5 * 60 * 1000; // items shown instantly from cache, then refreshed
const VIEWS_FRESH = 10 * 60 * 1000; // view counts change slowly: re-check search at most this often
const VIEWS_KEEP = 24 * 60 * 60 * 1000; // ...and keep showing the last known counts for up to a day

const SEARCH_BACKOFF = 5 * 60 * 1000;
let searchBackoffUntil = 0;
function searchAvailable(): boolean {
  return Date.now() >= searchBackoffUntil;
}
function pauseSearch(): void {
  searchBackoffUntil = Date.now() + SEARCH_BACKOFF;
}

interface IViewEntry {
  life: number;
  recent: number;
  at: number;
}
const STORE_PREFIX = 'pageRollup:v1:';
const memory: Record<string, { at: number; value: unknown }> = {};
const inflight: Record<string, Promise<unknown>> = {};

function readStore<T>(key: string, ttl: number): T | undefined {
  const hit = memory[key];
  if (hit && Date.now() - hit.at < ttl) return hit.value as T;
  try {
    const raw = window.sessionStorage.getItem(STORE_PREFIX + key);
    if (raw) {
      const parsed = JSON.parse(raw) as { at: number; value: T };
      if (Date.now() - parsed.at < ttl) {
        memory[key] = parsed;
        return parsed.value;
      }
    }
  } catch {
    // storage blocked or full: memory cache only
  }
  return undefined;
}

function writeStore(key: string, value: unknown): void {
  const entry = { at: Date.now(), value };
  memory[key] = entry;
  try {
    window.sessionStorage.setItem(STORE_PREFIX + key, JSON.stringify(entry));
  } catch {
    // ignore
  }
}

/** Runs the loader once per key at a time and caches the result. */
function cached<T>(key: string, ttl: number, load: () => Promise<T>): Promise<T> {
  const hit = readStore<T>(key, ttl);
  if (hit !== undefined) return Promise.resolve(hit);
  if (!inflight[key]) {
    inflight[key] = load().then(
      (v) => {
        delete inflight[key];
        writeStore(key, v);
        return v;
      },
      (e) => {
        delete inflight[key];
        throw e;
      }
    );
  }
  return inflight[key] as Promise<T>;
}

export class RollupService {

  constructor(private _http: SPHttpClient, private _currentWebUrl: string) {}

  public resolveSiteUrl(siteUrl: string): string {
    const url = (siteUrl || '').trim().replace(/\/+$/, '');
    return url || this._currentWebUrl.replace(/\/+$/, '');
  }

  public async getLists(siteUrl: string): Promise<IRollupList[]> {
    const web = this.resolveSiteUrl(siteUrl);
    return cached(`lists|${web}`, META_TTL, async () => {
      const data = await this._get(
        `${web}/_api/web/lists?$filter=Hidden eq false and BaseType eq 1&$select=Id,Title,BaseTemplate&$orderby=Title`
      );
      return (data.value || []).map((l: any) => ({ id: l.Id, title: l.Title, baseTemplate: l.BaseTemplate })) as IRollupList[];
    });
  }

  public async getFields(siteUrl: string, listId: string): Promise<IRollupField[]> {
    const web = this.resolveSiteUrl(siteUrl);
    return cached(`fields|${web}|${listId}`, META_TTL, () => this._loadFields(web, listId));
  }

  private async _loadFields(web: string, listId: string): Promise<IRollupField[]> {
    const data = await this._get(
      `${web}/_api/web/lists(guid'${listId}')/fields?$select=InternalName,Title,TypeAsString,Hidden`
    );
    const seen: Record<string, boolean> = {};
    const fields: IRollupField[] = (data.value || [])
      .filter((f: any) => SUPPORTED_FIELD_TYPES.indexOf(f.TypeAsString) !== -1)
      .filter((f: any) => !f.Hidden || USEFUL_HIDDEN_FIELDS.indexOf(f.InternalName) !== -1)
      .filter((f: any) => (f.InternalName.indexOf('_') !== 0 || USEFUL_HIDDEN_FIELDS.indexOf(f.InternalName) !== -1) && f.InternalName !== 'ContentType')
      .filter((f: any) => (seen[f.InternalName] ? false : (seen[f.InternalName] = true)))
      .map((f: any) => ({ internalName: f.InternalName, title: f.Title, type: f.TypeAsString }))
      .sort((a: IRollupField, b: IRollupField) => a.title.localeCompare(b.title));
    return fields;
  }

  /** Items from the last fetch of this exact query, if recent, with how old they are (ms). */
  public peekResult(query: IRollupQuery): { result: IRollupResult; age: number } | undefined {
    const key = this._resultKey(query);
    const hit = readStore<IRollupResult>(key, RESULT_TTL);
    return hit ? { result: hit, age: Date.now() - (memory[key] ? memory[key].at : Date.now()) } : undefined;
  }

  public storeResult(query: IRollupQuery, result: IRollupResult): void {
    writeStore(this._resultKey(query), result);
  }

  private _resultKey(query: IRollupQuery): string {
    return `items|${this.resolveSiteUrl(query.siteUrl)}|${JSON.stringify(query)}`;
  }

  /**
   * @param deferViews when true, view counts are not fetched here (unless sorting by views); the caller shows the
   *                   tiles first and then calls loadViews().
   */
  public async getItems(originalQuery: IRollupQuery, deferViews?: boolean): Promise<IRollupResult> {
    const web = this.resolveSiteUrl(originalQuery.siteUrl);
    const origin = new URL(web).origin;
    const fields = await this.getFields(web, originalQuery.listId);
    const has = (name: string): boolean => fields.some((f) => f.internalName === name);
    // Work out from the columns whether this is a Site Pages library (no extra request needed)
    const query: IRollupQuery = { ...originalQuery, isPagesLibrary: originalQuery.isPagesLibrary || (has('BannerImageUrl') && has('PromotedState')) };
    const ids = query.isPagesLibrary ? undefined : await this._getWebIds(web);

    const dateField = has(query.dateField) ? query.dateField : 'Created';
    const clientFilters = query.filters.filter((f) => f.field && has(f.field));
    const sortByLikes = query.sortBy === 'likes';
    // Both need every candidate before sorting, so they scan instead of stopping at the limit
    const sortByViews = query.sortBy === 'views' || sortByLikes;
    const likeField = LIKE_FIELDS.filter(has)[0] || '';
    const commentField = COMMENT_FIELDS.filter(has)[0] || '';

    const viewFields = this._unique([
      ...ALWAYS_SELECT,
      dateField,
      ...(query.isPagesLibrary ? ['BannerImageUrl', 'Description'] : []),
      ...(query.thumbnailField ? [query.thumbnailField] : []),
      ...query.displayFields,
      ...(query.detailsField ? [query.detailsField] : []),
      ...clientFilters.map((f) => f.field),
      ...(query.includeSocial || sortByLikes ? [likeField, commentField] : [])
    ]).filter((f) => ALWAYS_SELECT.indexOf(f) !== -1 || has(f));

    // Scanning (filters / most viewed / most liked) reads big batches to keep round trips down;
    // otherwise just enough rows for the limit, plus a few spare to skip page templates.
    const scanning = clientFilters.length > 0 || sortByViews;
    const viewXml = this._buildViewXml(query, dateField, viewFields, has, scanning ? 500 : query.maxItems + (query.isPagesLibrary ? 10 : 0));
    const maxScan = sortByViews ? 1000 : 2000;
    const target = query.maxItems;

    const matched: IRollupItem[] = [];
    let scanned = 0;
    let nextHref = '';
    let truncated = false;

    for (;;) {
      const page = await this._post(
        `${web}/_api/web/lists(guid'${query.listId}')/RenderListDataAsStream${nextHref}`,
        { parameters: { RenderOptions: 2, ViewXml: viewXml } }
      );
      const rows: any[] = page.Row || [];
      scanned += rows.length;

      for (const row of rows) {
        if (query.isPagesLibrary && /\/Templates$/i.test(row.FileDirRef || '')) {
          continue; // page templates folder
        }
        if (!rowMatches(row, clientFilters, query.filterMatch)) {
          continue;
        }
        const item = this._toItem(row, web, origin, query, viewFields, ids);
        if (query.includeSocial || sortByLikes) {
          if (likeField) item.likes = Number(row[likeField]) || 0;
          if (commentField) item.comments = Number(row[commentField]) || 0;
        }
        matched.push(item);
        if (!sortByViews && matched.length >= target) {
          break;
        }
      }

      if ((!sortByViews && matched.length >= target) || !page.NextHref) {
        break;
      }
      if (scanned >= maxScan) {
        truncated = true;
        break;
      }
      nextHref = page.NextHref;
    }

    let items = matched;
    if (sortByLikes) {
      // Trim before looking up views so the search query only names the items shown
      items = items.slice().sort((a, b) => (b.likes || 0) - (a.likes || 0)).slice(0, target);
    }
    const sortingByViewCount = sortByViews && !sortByLikes;
    if (query.includeViews && deferViews && !sortingByViewCount) {
      return { items: items.slice(0, target), scanned, truncated, viewsPending: true };
    }
    if (query.includeViews || sortingByViewCount) {
      await this._attachViews(web, query.listId, items, sortingByViewCount);
      if (sortByViews && !sortByLikes) {
        const pick = (i: IRollupItem): number => (query.viewsPeriod === 'recent' ? i.viewsRecent : i.viewsLifetime) || 0;
        items = items.slice().sort((a, b) => pick(b) - pick(a));
      }
    }

    return { items: items.slice(0, target), scanned, truncated };
  }

  /** Adds view counts to items already on screen (returns new objects). */
  /**
   * Marks items published (or created/modified, per the date column) within the last query.newDays days.
   * Exact dates come from the REST items endpoint, because RenderListDataAsStream formats dates for the
   * viewer's region (e.g. day/month vs month/day), which can't be parsed reliably.
   */
  public async loadNewFlags(query: IRollupQuery, items: IRollupItem[]): Promise<IRollupItem[]> {
    if (!query.newDays || items.length === 0) return items;
    const web = this.resolveSiteUrl(query.siteUrl);
    const fields = await this.getFields(web, query.listId);
    const has = (name: string): boolean => fields.some((f) => f.internalName === name);
    const dateField = has(query.dateField) ? query.dateField : 'Created';
    const filter = items.map((i) => `Id eq ${i.id}`).join(' or ');
    const data = await this._get(
      `${web}/_api/web/lists(guid'${query.listId.replace(/[{}]/g, '')}')/items?$select=Id,${dateField}&$filter=${encodeURIComponent(filter)}&$top=${items.length}`
    );
    const cutoff = Date.now() - query.newDays * 24 * 60 * 60 * 1000;
    const isNew: Record<number, boolean> = {};
    (data.value || []).forEach((r: any) => {
      const t = r[dateField] ? Date.parse(r[dateField]) : NaN;
      isNew[Number(r.Id)] = !isNaN(t) && t >= cutoff && t <= Date.now() + 60000;
    });
    return items.map((i) => ({ ...i, isNew: !!isNew[i.id] }));
  }

  /** Items with the last known view counts filled in (instant, no request). */
  public knownViews(query: IRollupQuery, items: IRollupItem[]): IRollupItem[] {
    const known = this._viewCache(query);
    const now = Date.now();
    return items.map((i) => {
      const k = known[i.id];
      return k && now - k.at < VIEWS_KEEP && i.viewsLifetime === undefined ? { ...i, viewsLifetime: k.life, viewsRecent: k.recent } : i;
    });
  }

  /**
   * Fills in view counts from SharePoint search for items whose counts are missing or more than 10 minutes old.
   * Never throws: if search fails (it can be slow or throttled) it retries once, then keeps the last known counts.
   * `complete` is false when some counts could not be refreshed, so the caller can try again next time.
   */
  public async loadViews(query: IRollupQuery, items: IRollupItem[]): Promise<{ items: IRollupItem[]; complete: boolean }> {
    const web = this.resolveSiteUrl(query.siteUrl);
    const known = this._viewCache(query);
    const now = Date.now();
    const stale = items.filter((i) => !known[i.id] || now - known[i.id].at >= VIEWS_FRESH);
    let complete = true;

    if (stale.length > 0) {
      const copy = stale.map((i) => ({ ...i, viewsLifetime: undefined, viewsRecent: undefined }));
      let ok = false;
      if (searchAvailable()) {
        ok = await this._attachViews(web, query.listId, copy, false);
        if (!ok) {
          // Search sometimes rejects the item-by-item query (500). Try once more with the simplest possible
          // query: every item in the library, matched up here.
          await new Promise((resolve) => window.setTimeout(resolve, 1000));
          ok = await this._attachViews(web, query.listId, copy, true);
        }
        if (!ok) {
          // Search is having a bad time: stop asking for a few minutes (fewer errors, less load on the page)
          pauseSearch();
        }
      }
      if (ok) {
        copy.forEach((i) => {
          if (i.viewsLifetime !== undefined) {
            known[i.id] = { life: i.viewsLifetime || 0, recent: i.viewsRecent || 0, at: Date.now() };
          }
        });
        writeStore(this._viewKey(query), known);
      }
      complete = ok && copy.every((i) => i.viewsLifetime !== undefined);
    }

    const result = items.map((i) => {
      const k = known[i.id];
      return k && now - k.at < VIEWS_KEEP ? { ...i, viewsLifetime: k.life, viewsRecent: k.recent } : { ...i };
    });
    return { items: result, complete };
  }

  private _viewKey(query: IRollupQuery): string {
    return `views|${this.resolveSiteUrl(query.siteUrl)}|${query.listId}`;
  }

  private _viewCache(query: IRollupQuery): Record<number, IViewEntry> {
    return readStore<Record<number, IViewEntry>>(this._viewKey(query), VIEWS_KEEP) || {};
  }

  // ---------------------------------------------------------------------------

  private _buildViewXml(
    query: IRollupQuery,
    dateField: string,
    viewFields: string[],
    has: (name: string) => boolean,
    rowLimit: number
  ): string {
    const conditions: string[] = [`<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq>`];

    const window = getDateWindow(query.dateRange, query.lastDays);
    if (window.start) {
      conditions.push(`<Geq><FieldRef Name='${dateField}'/>${this._dateValue(window.start)}</Geq>`);
    }
    if (window.end) {
      conditions.push(`<Lt><FieldRef Name='${dateField}'/>${this._dateValue(window.end)}</Lt>`);
    }
    if (query.isPagesLibrary && query.publishedOnly && has('FirstPublishedDate')) {
      conditions.push(`<IsNotNull><FieldRef Name='FirstPublishedDate'/></IsNotNull>`);
    }
    if (query.isPagesLibrary && query.newsOnly && has('PromotedState')) {
      conditions.push(`<Eq><FieldRef Name='PromotedState'/><Value Type='Number'>2</Value></Eq>`);
    }

    const where = conditions.reduce((acc, c) => (acc ? `<And>${acc}${c}</And>` : c), '');
    const orderBy =
      query.sortBy === 'title'
        ? `<FieldRef Name='Title' Ascending='TRUE'/>`
        : `<FieldRef Name='${dateField}' Ascending='${query.sortBy === 'dateAsc' ? 'TRUE' : 'FALSE'}'/>`;

    return (
      `<View Scope='RecursiveAll'>` +
      `<Query><Where>${where}</Where><OrderBy>${orderBy}</OrderBy></Query>` +
      `<ViewFields>${viewFields.map((f) => `<FieldRef Name='${f}'/>`).join('')}</ViewFields>` +
      `<RowLimit Paged='TRUE'>${Math.max(1, Math.min(rowLimit, 500))}</RowLimit>` +
      `</View>`
    );
  }

  private _dateValue(d: Date): string {
    return `<Value Type='DateTime' IncludeTimeValue='TRUE' StorageTZ='TRUE'>${d.toISOString()}</Value>`;
  }

  private _toItem(row: any, web: string, origin: string, query: IRollupQuery, viewFields: string[], ids?: IWebIds): IRollupItem {
    const fileRef: string = row.FileRef || '';
    const fileType = String(row.File_x0020_Type || fileRef.split('.').pop() || '').toLowerCase();
    const absoluteFile = `${origin}${fileRef}`;
    const url = OFFICE_TYPES.indexOf(fileType) !== -1 ? `${absoluteFile}?web=1` : absoluteFile;

    const values: Record<string, string> = {};
    viewFields.forEach((f) => {
      values[f] = getFieldText(row, f);
    });

    const fileName: string = row.FileLeafRef || '';
    const title = (row.Title && String(row.Title).trim()) || fileName.replace(/\.[^.]+$/, '');
    values.Title = title;

    const image = this._resolveImage(row, web, origin, absoluteFile, fileType, query, ids);

    return {
      id: Number(row.ID),
      url,
      title,
      fileType,
      originalImage: image.original,
      previewBase: image.previewBase,
      thumbTemplate: image.thumbTemplate,
      extraSources: image.extraSources,
      isDocument: image.isDocument,
      details: query.detailsField && viewFields.indexOf(query.detailsField) !== -1 ? getFieldTextMultiline(row, query.detailsField) : '',
      values
    };
  }

  private _resolveImage(
    row: any,
    web: string,
    origin: string,
    absoluteFile: string,
    fileType: string,
    query: IRollupQuery,
    ids?: IWebIds
  ): { original: string; previewBase: string; thumbTemplate: string; extraSources: string[]; isDocument: boolean } {
    const candidates = [query.thumbnailField, query.isPagesLibrary ? 'BannerImageUrl' : ''].filter((f) => !!f);
    for (const field of candidates) {
      const raw = this._imageFromField(row, field);
      if (raw && raw.toLowerCase().indexOf(DEFAULT_PAGE_IMAGE) === -1) {
        return { ...this._describeImage(this._absolute(raw, origin), web, origin), extraSources: [], isDocument: false };
      }
    }
    if (IMAGE_TYPES.indexOf(fileType) !== -1) {
      // The item is itself a picture (e.g. a picture library)
      return { ...this._describeImage(absoluteFile, web, origin), extraSources: [], isDocument: false };
    }
    if (PREVIEWABLE.indexOf(fileType) !== -1) {
      return { ...this._documentPreview(row, web, absoluteFile, query.listId, ids), original: '', isDocument: true };
    }
    return { original: '', previewBase: '', thumbTemplate: '', extraSources: [], isDocument: true };
  }

  /**
   * Document thumbnails, best first:
   *  1. the thumbnail API modern SharePoint uses for its own library tiles (needs site/web ids)
   *  2. the preview handler addressed by site/web/file ids
   *  3. the preview handler addressed by path
   * The tile tries each in turn if one fails, then falls back to a file-type placeholder.
   */
  private _documentPreview(
    row: any,
    web: string,
    absoluteFile: string,
    listId: string,
    ids?: IWebIds
  ): { previewBase: string; thumbTemplate: string; extraSources: string[] } {
    const uniqueId = String(row.UniqueId || '').replace(/[{}]/g, '');
    const clean = (g: string): string => g.replace(/[{}]/g, '');
    const thumbTemplate = ids
      ? `${web}/_api/v2.1/sites/${ids.host},${clean(ids.siteId)},${clean(ids.webId)}/lists/${clean(listId)}/items/${row.ID}/driveItem/thumbnails/0/c{W}x99999/content?prefer=noRedirect`
      : '';
    const byIds =
      ids && uniqueId
        ? `${web}/_layouts/15/getpreview.ashx?guidSite=${clean(ids.siteId)}&guidWeb=${clean(ids.webId)}&guidFile=${uniqueId}`
        : '';
    const byPath = this._previewBase(web, absoluteFile);
    return {
      thumbTemplate,
      previewBase: byIds || byPath,
      extraSources: byIds ? [`${byPath}&resolution=3`] : []
    };
  }

  /** Site and web ids for a site address (cached), needed for document thumbnails. */
  private _getWebIds(web: string): Promise<IWebIds | undefined> {
    return cached(`ids|${web}`, META_TTL, () =>
      Promise.all([this._get(`${web}/_api/site?$select=Id`), this._get(`${web}/_api/web?$select=Id`)]).then(
        ([site, w]) => ({ host: new URL(web).hostname, siteId: String(site.Id), webId: String(w.Id) }) as IWebIds
      )
    ).catch(() => undefined);
  }


  /**
   * Banner images are often stored as an already-shrunk preview or thumbnail URL, which looks blurry on big tiles.
   * Unwrap those to the real image, and keep a sized-thumbnail template where SharePoint gave us one.
   */
  private _describeImage(abs: string, web: string, origin: string): { original: string; previewBase: string; thumbTemplate: string } {
    let original = abs;
    let thumbTemplate = '';

    // .../_layouts/15/getpreview.ashx?path=<file>&resolution=n  ->  <file>
    const previewMatch = /getpreview\.ashx\?(?:.*&)?path=([^&]+)/i.exec(abs);
    if (previewMatch) {
      try {
        original = this._absolute(decodeURIComponent(previewMatch[1]), origin);
      } catch {
        original = abs;
      }
    }

    // .../thumbnails/0/c400x99999/content  ->  ask for the size we need
    if (/\/thumbnails\/\d+\/c\d+x\d+\/content/i.test(abs)) {
      thumbTemplate = abs.replace(/\/c\d+x\d+\/content/i, '/c{W}x99999/content');
    }

    const sameTenant = original.indexOf(origin) === 0 && original.indexOf('/_layouts/') === -1 && original.indexOf('/_api/') === -1;
    return { original, previewBase: sameTenant ? this._previewBase(web, original) : '', thumbTemplate };
  }

  private _previewBase(web: string, absFile: string): string {
    // Banner addresses are often already URL-encoded (e.g. "%2C" for a comma). Decode first so the path
    // isn't encoded twice, which makes getpreview.ashx answer 400 Bad Request.
    // decodeURIComponent, not decodeURI: decodeURI leaves %2C (comma) and %2F (slash) encoded.
    let path = absFile;
    try {
      path = decodeURIComponent(absFile);
    } catch {
      // malformed escape: use as-is
    }
    return `${web}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(path)}`;
  }

  private _absolute(raw: string, origin: string): string {
    return raw.indexOf('http') === 0 ? raw : `${origin}${raw.indexOf('/') === 0 ? '' : '/'}${raw}`;
  }

  private _imageFromField(row: any, field: string): string {
    let raw = row[field];
    if (!raw) {
      return '';
    }
    if (typeof raw === 'string' && raw.trim().indexOf('{') === 0) {
      try {
        raw = JSON.parse(raw);
      } catch {
        return '';
      }
    }
    if (typeof raw === 'object') {
      return raw.serverUrl && raw.serverRelativeUrl
        ? `${raw.serverUrl}${raw.serverRelativeUrl}`
        : raw.serverRelativeUrl || raw.Url || raw.url || '';
    }
    const text = String(raw);
    // "url, description" form
    const comma = text.indexOf(', ');
    return comma > 0 && text.indexOf('http') === 0 ? text.substring(0, comma) : text;
  }

  /** Page views come from SharePoint search (ViewsLifeTime / ViewsRecent); no Graph permission needed. */
  private async _attachViews(web: string, listId: string, items: IRollupItem[], wholeList: boolean): Promise<boolean> {
    if (items.length === 0) {
      return true;
    }
    const cleanListId = listId.replace(/[{}]/g, '').toLowerCase();
    let queryText = `ListId:${cleanListId}`;
    if (!wholeList) {
      queryText += ` AND (${items.map((i) => `ListItemID:${i.id}`).join(' OR ')})`;
    }
    const url =
      `${web}/_api/search/query?querytext='${encodeURIComponent(queryText)}'` +
      `&selectproperties='${encodeURIComponent('ListItemID,ViewsLifeTime,ViewsRecent,Path')}'` +
      `&rowlimit=500&trimduplicates=false` +
      `&sourceid='8413cd39-2156-4e00-b54d-11efd9abdb89'`;

    try {
      const data = await this._get(url);
      const rows: any[] =
        (data.PrimaryQueryResult && data.PrimaryQueryResult.RelevantResults &&
          data.PrimaryQueryResult.RelevantResults.Table && data.PrimaryQueryResult.RelevantResults.Table.Rows) || [];
      const byId: Record<string, { life: number; recent: number }> = {};
      const byPath: Record<string, { life: number; recent: number }> = {};
      rows.forEach((r: any) => {
        const cells: Record<string, string> = {};
        (r.Cells || []).forEach((c: any) => {
          cells[c.Key] = c.Value;
        });
        const entry = { life: Number(cells.ViewsLifeTime) || 0, recent: Number(cells.ViewsRecent) || 0 };
        if (cells.ListItemID) {
          byId[cells.ListItemID] = entry;
        }
        if (cells.Path) {
          byPath[this._normaliseUrl(cells.Path)] = entry;
        }
      });
      items.forEach((i) => {
        const hit = byId[String(i.id)] || (i.url ? byPath[this._normaliseUrl(i.url)] : undefined);
        // Item-by-item query: not found means search has no views for it (0).
        // Whole-library query (max 500 results): not found may just mean it wasn't in the first 500.
        i.viewsLifetime = hit ? hit.life : wholeList ? i.viewsLifetime : 0;
        i.viewsRecent = hit ? hit.recent : wholeList ? i.viewsRecent : 0;
      });
      return true;
    } catch (e) {
      // Search can be slow, throttled or briefly unavailable; the caller keeps the last known counts.
      console.warn('[PageRollup] Could not load view counts', e);
      return false;
    }
  }

  private _normaliseUrl(u: string): string {
    try {
      return decodeURIComponent(u.split('?')[0]).toLowerCase();
    } catch {
      return u.split('?')[0].toLowerCase();
    }
  }

  private _unique(values: string[]): string[] {
    return values.filter((v, i) => !!v && values.indexOf(v) === i);
  }

  private async _get(url: string): Promise<any> {
    const res = await this._http.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' }
    });
    return this._json(res);
  }

  private async _post(url: string, body: any): Promise<any> {
    const res = await this._http.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify(body)
    });
    return this._json(res);
  }

  private async _json(res: SPHttpClientResponse): Promise<any> {
    if (!res.ok) {
      let message = `${res.status} ${res.statusText}`;
      try {
        const err = await res.json();
        message = (err && (err['odata.error'] ? err['odata.error'].message.value : err.error && err.error.message)) || message;
      } catch {
        // ignore parse errors
      }
      throw new Error(message);
    }
    return res.json();
  }
}
