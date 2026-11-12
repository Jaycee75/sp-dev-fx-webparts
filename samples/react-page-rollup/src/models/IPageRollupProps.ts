export type RollupLayout = 'cards' | 'tiles' | 'hero' | 'mosaic' | 'filmstrip' | 'list' | 'spotlight' | 'social' | 'ticker';
export type CarouselBackground = 'none' | 'theme' | 'blur' | 'image' | 'color' | 'gradient';
export type ImageQuality = 'sharp' | 'fast' | 'original';
export type ImageFit = 'cover' | 'contain';
export type TitleSize = 'small' | 'medium' | 'large' | 'xlarge';
export type TitleAlign = 'left' | 'center' | 'right';
export type HoverEffect = 'zoom' | 'lift' | 'reveal' | 'none';
export type OverlayStyle = 'gradient' | 'glass' | 'solid';
export type DateRange = 'all' | 'lastDays' | 'thisQuarter' | 'lastQuarter' | 'thisYear';
export type SortBy = 'dateDesc' | 'dateAsc' | 'title' | 'views' | 'likes';
export type FilterOperator = 'eq' | 'neq' | 'contains' | 'empty' | 'notEmpty';
export type SizeMode = 'auto' | 'fixed';
export type AspectRatio = '16:9' | '4:3' | '1:1' | '3:4';
export type ViewsPeriod = 'lifetime' | 'recent';

/** Every setting stored by the web part (flat, so the built-in property pane can edit it). */
export interface IPageRollupWebPartProps {
  title: string;
  titleSize: TitleSize;
  titleAlign: TitleAlign;

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
  detailsField: string;
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
  showLikes: boolean;
  showComments: boolean;
  openInNewTab: boolean;
  imageQuality: ImageQuality;
  imageFit: ImageFit;
  /** Set by the toolbox entry: 'documents' starts on the Documents library instead of Site Pages. */
  startWith: string;

  // Carousels (spotlight, social cards, slide-up ticker)
  carouselAutoplay: boolean;
  carouselInterval: number;
  carouselSpeed: number;
  carouselPauseOnHover: boolean;
  carouselLoop: boolean;
  carouselArrows: boolean;
  carouselDots: boolean;
  carouselSlideWidth: number;
  carouselBackground: CarouselBackground;
  carouselBgColor: string;
  carouselBgColor2: string;
  carouselBgImage: string;
  carouselBgImageBlur: boolean;
  carouselBgDim: number;
  tickerRows: number;

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
  includeSocial: boolean;
  viewsPeriod: ViewsPeriod;
  detailsField: string;
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
  /** Full-size picture (banner image, picture column, or the file itself for images). */
  originalImage: string;
  /** SharePoint preview handler URL without a resolution; used for documents and the "fast" quality option. */
  previewBase: string;
  /** Sized thumbnail URL with a {W} placeholder for the width in pixels, when the source is a thumbnail API URL. */
  thumbTemplate: string;
  /** More picture sources to try, in order, if the ones above fail (document previews). */
  extraSources: string[];
  /** True for files (Word, PDF, ...) rather than pages or pictures; previews show the top of the page. */
  isDocument: boolean;
  /** Display text per internal field name. */
  values: Record<string, string>;
  /** Full text of the details column, line breaks kept (shown on hover). */
  details?: string;
  /** Likes and comments: from the page's _LikeCount/_CommentCount, or a library's ratings 'LikesCount' column. Undefined when the library doesn't track them. */
  likes?: number;
  comments?: number;
  viewsLifetime?: number;
  viewsRecent?: number;
}
