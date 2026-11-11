export type RollupLayout = 'cards' | 'tiles' | 'hero' | 'mosaic' | 'filmstrip' | 'list';
export type HoverEffect = 'zoom' | 'lift' | 'reveal' | 'none';
export type OverlayStyle = 'gradient' | 'glass' | 'solid';
export type DateRange = 'all' | 'lastDays' | 'thisQuarter' | 'lastQuarter' | 'thisYear';
export type SortBy = 'dateDesc' | 'dateAsc' | 'title' | 'views';
export type FilterOperator = 'eq' | 'neq' | 'contains' | 'empty' | 'notEmpty';
export type SizeMode = 'auto' | 'fixed';
export type AspectRatio = '16:9' | '4:3' | '1:1' | '3:4';
export type ViewsPeriod = 'lifetime' | 'recent';

/** Every setting stored by the web part (flat, so the built-in property pane can edit it). */
export interface IPageRollupWebPartProps {
  title: string;

  // Source
  siteUrl: string;
  listId: string;
  publishedOnly: boolean;
  newsOnly: boolean;

  // Filters (up to 3)
  filterMatch: 'all' | 'any';
  filter1Field: string;
  filter1Operator: FilterOperator;
  filter1Value: string;
  filter2Field: string;
  filter2Operator: FilterOperator;
  filter2Value: string;
  filter3Field: string;
  filter3Operator: FilterOperator;
  filter3Value: string;

  // Date window, sort, limit
  dateField: string;
  dateRange: DateRange;
  lastDays: number;
  sortBy: SortBy;
  maxItems: number;

  // Display
  layout: RollupLayout;
  hoverEffect: HoverEffect;
  showOverlay: boolean;
  overlayStyle: OverlayStyle;
  overlayField1: string;
  overlayField2: string;
  overlayField3: string;
  overlayField4: string;
  overlayField5: string;
  showFieldLabels: boolean;
  thumbnailField: string;
  showViews: boolean;
  viewsPeriod: ViewsPeriod;
  openInNewTab: boolean;

  // Size
  sizeMode: SizeMode;
  fixedHeight: number;
  columns: number;
  tileSize: number;
  aspectRatio: AspectRatio;
  gap: number;
  cornerRadius: number;
}

export interface IRollupFilter {
  field: string;
  operator: FilterOperator;
  values: string[];
}

/** The subset of settings that decides which items are fetched. */
export interface IRollupQuery {
  siteUrl: string;
  listId: string;
  isPagesLibrary: boolean;
  publishedOnly: boolean;
  newsOnly: boolean;
  filters: IRollupFilter[];
  filterMatch: 'all' | 'any';
  dateField: string;
  dateRange: DateRange;
  lastDays: number;
  sortBy: SortBy;
  maxItems: number;
  displayFields: string[];
  thumbnailField: string;
  includeViews: boolean;
  viewsPeriod: ViewsPeriod;
}

export interface IRollupField {
  internalName: string;
  title: string;
  type: string;
}

export interface IRollupList {
  id: string;
  title: string;
  baseTemplate: number;
}

export interface IRollupItem {
  id: number;
  url: string;
  title: string;
  fileType: string;
  imageUrl: string;
  fallbackImageUrl: string;
  /** Display text per internal field name. */
  values: Record<string, string>;
  viewsLifetime?: number;
  viewsRecent?: number;
}
