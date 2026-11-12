import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IRollupField, IRollupItem, IRollupList, IRollupQuery } from '../models/IPageRollupProps';
import { getDateWindow } from '../utils/dateRange';
import { getFieldText, getFieldTextMultiline, rowMatches } from '../utils/fieldValues';

/* eslint-disable @typescript-eslint/no-explicit-any */

/** Site Pages libraries use this list template. */
export const SITE_PAGES_TEMPLATE = 119;

/** Columns that are hidden but still useful on Site Pages libraries. */
const USEFUL_HIDDEN_FIELDS = ['FirstPublishedDate', 'Description', 'BannerImageUrl', 'PromotedState', 'FileLeafRef', 'Modified', 'Created'];

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
}

export class RollupService {
  private _fieldCache: Record<string, IRollupField[]> = {};
  private _listCache: Record<string, IRollupList[]> = {};
  private _idCache: Record<string, Promise<IWebIds | undefined>> = {};

  constructor(private _http: SPHttpClient, private _currentWebUrl: string) {}

  public resolveSiteUrl(siteUrl: string): string {
    const url = (siteUrl || '').trim().replace(/\/+$/, '');
    return url || this._currentWebUrl.replace(/\/+$/, '');
  }

  public async getLists(siteUrl: string): Promise<IRollupList[]> {
    const web = this.resolveSiteUrl(siteUrl);
    if (this._listCache[web]) {
      return this._listCache[web];
    }
    const data = await this._get(
      `${web}/_api/web/lists?$filter=Hidden eq false and BaseType eq 1&$select=Id,Title,BaseTemplate&$orderby=Title`
    );
    const lists: IRollupList[] = (data.value || []).map((l: any) => ({
      id: l.Id,
      title: l.Title,
      baseTemplate: l.BaseTemplate
    }));
    this._listCache[web] = lists;
    return lists;
  }

  public async getFields(siteUrl: string, listId: string): Promise<IRollupField[]> {
    const web = this.resolveSiteUrl(siteUrl);
    const key = `${web}|${listId}`;
    if (this._fieldCache[key]) {
      return this._fieldCache[key];
    }
    const data = await this._get(
      `${web}/_api/web/lists(guid'${listId}')/fields?$select=InternalName,Title,TypeAsString,Hidden`
    );
    const seen: Record<string, boolean> = {};
    const fields: IRollupField[] = (data.value || [])
      .filter((f: any) => SUPPORTED_FIELD_TYPES.indexOf(f.TypeAsString) !== -1)
      .filter((f: any) => !f.Hidden || USEFUL_HIDDEN_FIELDS.indexOf(f.InternalName) !== -1)
      .filter((f: any) => f.InternalName.indexOf('_') !== 0 && f.InternalName !== 'ContentType')
      .filter((f: any) => (seen[f.InternalName] ? false : (seen[f.InternalName] = true)))
      .map((f: any) => ({ internalName: f.InternalName, title: f.Title, type: f.TypeAsString }))
      .sort((a: IRollupField, b: IRollupField) => a.title.localeCompare(b.title));
    this._fieldCache[key] = fields;
    return fields;
  }

  public async getItems(query: IRollupQuery): Promise<IRollupResult> {
    const web = this.resolveSiteUrl(query.siteUrl);
    const origin = new URL(web).origin;
    const fields = await this.getFields(web, query.listId);
    const ids = query.isPagesLibrary ? undefined : await this._getWebIds(web);
    const has = (name: string): boolean => fields.some((f) => f.internalName === name);

    const dateField = has(query.dateField) ? query.dateField : 'Created';
    const clientFilters = query.filters.filter((f) => f.field && has(f.field));
    const needsClientFiltering = clientFilters.length > 0 || query.isPagesLibrary;
    const sortByViews = query.sortBy === 'views';

    const viewFields = this._unique([
      ...ALWAYS_SELECT,
      dateField,
      ...(query.isPagesLibrary ? ['BannerImageUrl', 'Description'] : []),
      ...(query.thumbnailField ? [query.thumbnailField] : []),
      ...query.displayFields,
      ...(query.detailsField ? [query.detailsField] : []),
      ...clientFilters.map((f) => f.field)
    ]).filter((f) => ALWAYS_SELECT.indexOf(f) !== -1 || has(f));

    const viewXml = this._buildViewXml(query, dateField, viewFields, has, needsClientFiltering || sortByViews ? 100 : query.maxItems);
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
        matched.push(this._toItem(row, web, origin, query, viewFields, ids));
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
    if (query.includeViews || sortByViews) {
      await this._attachViews(web, query.listId, items, sortByViews);
      if (sortByViews) {
        const pick = (i: IRollupItem): number => (query.viewsPeriod === 'recent' ? i.viewsRecent : i.viewsLifetime) || 0;
        items = items.slice().sort((a, b) => pick(b) - pick(a));
      }
    }

    return { items: items.slice(0, target), scanned, truncated };
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
    if (!this._idCache[web]) {
      this._idCache[web] = Promise.all([this._get(`${web}/_api/site?$select=Id`), this._get(`${web}/_api/web?$select=Id`)])
        .then(([site, w]) => ({ host: new URL(web).hostname, siteId: String(site.Id), webId: String(w.Id) }))
        .catch(() => undefined);
    }
    return this._idCache[web];
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
    return `${web}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(absFile)}`;
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
  private async _attachViews(web: string, listId: string, items: IRollupItem[], wholeList: boolean): Promise<void> {
    if (items.length === 0) {
      return;
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
        const hit = byId[String(i.id)] || byPath[this._normaliseUrl(i.url)];
        i.viewsLifetime = hit ? hit.life : 0;
        i.viewsRecent = hit ? hit.recent : 0;
      });
    } catch (e) {
      // Search can be unavailable (e.g. not crawled yet); just hide the view badges.
      console.warn('[PageRollup] Could not load view counts', e);
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
