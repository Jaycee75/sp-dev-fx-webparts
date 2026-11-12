import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { DisplayMode } from '@microsoft/sp-core-library';
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  IPropertyPaneField,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneHorizontalRule,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PageRollupWebPartStrings';
import { PageRollup, IPageRollupProps } from '../../components/PageRollup';
import { IRollupDisplay } from '../../components/RollupTile';
import {
  FilterOperator,
  IPageRollupWebPartProps,
  IRollupField,
  IRollupList,
  IRollupQuery
} from '../../models/IPageRollupProps';
import { RollupService, SITE_PAGES_TEMPLATE } from '../../services/RollupService';
import { describeDateWindow } from '../../utils/dateRange';
import { parseFilterValues } from '../../utils/fieldValues';

const NONE = '';
const DEFAULTS: Partial<IPageRollupWebPartProps> = {
  title: '',
  titleSize: 'medium',
  titleAlign: 'left',
  siteUrl: '',
  listId: '',
  publishedOnly: true,
  newsOnly: false,
  filterMatch: 'all',
  filter1Operator: 'eq',
  filter2Operator: 'eq',
  filter3Operator: 'eq',
  dateField: 'FirstPublishedDate',
  dateRange: 'all',
  lastDays: 30,
  sortBy: 'dateDesc',
  maxItems: 12,
  layout: 'tiles',
  hoverEffect: 'zoom',
  detailsField: 'Description',
  showOverlay: true,
  overlayStyle: 'gradient',
  overlayField1: 'Title',
  overlayField2: 'Description',
  overlayField3: '',
  overlayField4: '',
  overlayField5: '',
  showFieldLabels: false,
  thumbnailField: '',
  showViews: true,
  viewsPeriod: 'lifetime',
  showLikes: true,
  showComments: false,
  highlightNew: false,
  newDays: 7,
  newLabel: 'New',
  newGlow: true,
  openInNewTab: false,
  imageQuality: 'sharp',
  imageFit: 'cover',
  startWith: '',
  carouselAutoplay: true,
  carouselInterval: 5,
  carouselSpeed: 700,
  carouselPauseOnHover: true,
  carouselLoop: true,
  carouselArrows: true,
  carouselDots: true,
  carouselSlideWidth: 60,
  carouselBackground: 'blur',
  carouselBgColor: '#1b1f3b',
  carouselBgColor2: '#5b2a86',
  carouselBgImage: '',
  carouselBgImageBlur: true,
  carouselBgDim: 40,
  tickerRows: 3,
  sizeMode: 'auto',
  fixedHeight: 480,
  columns: 0,
  tileSize: 260,
  aspectRatio: '16:9',
  gap: 16,
  cornerRadius: 12
};

export default class PageRollupWebPart extends BaseClientSideWebPart<IPageRollupWebPartProps> {
  private _service!: RollupService;
  private _theme: IReadonlyTheme | undefined;
  private _lists: IRollupList[] = [];
  private _fields: IRollupField[] = [];
  private _loadingLists = false;
  private _loadingFields = false;
  private _sourceError = '';

  /** Library chosen automatically (Site Pages / Documents) for a new web part; kept in memory, not saved. */
  private _autoListId = '';
  private _booting = false;

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Don't write defaults into this.properties here: in edit mode SharePoint treats any property change
    // as an edit and saves a draft. Defaults are applied when reading settings (see _settings()).

    this._service = new RollupService(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
    try {
      if (window.location.search.indexOf('rollupdebug=1') !== -1) {
        console.info(`[PageRollup] version ${this.manifest.version} (instance ${this.instanceId})`);
      }
    } catch {
      // ignore
    }

    const themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._theme = themeProvider.tryGetTheme();
    themeProvider.themeChangedEvent.add(this, (args: ThemeChangedEventArgs) => {
      this._theme = args.theme;
      this.render();
    });

    // Nothing is awaited here, so the web part draws its loading state straight away.
    // SharePoint times onInit and flags web parts that are slow to start.
    this._bootstrap().catch(() => undefined);
  }

  /** Background start-up: pick a library for new web parts and load column names for labels. */
  private async _bootstrap(): Promise<void> {
    const listId = this.properties.listId;
    if (!listId) {
      this._booting = true;
      try {
        const lists = await this._service.getLists(this.properties.siteUrl);
        this._lists = lists;
        const lib = this._defaultLibrary(lists);
        this._autoListId = lib ? lib.id : '';
      } catch {
        // Leave unconfigured; the placeholder asks the editor to pick a library.
      }
      this._booting = false;
      this.render();
    }
    const effective = this.properties.listId || this._autoListId;
    if (effective) {
      this._fields = await this._service.getFields(this.properties.siteUrl, effective).catch(() => []);
      this.render(); // column titles for "Show column names"
    }
  }

  /** Saved settings with defaults filled in for anything not set (without changing the saved properties). */
  private _settings(): IPageRollupWebPartProps {
    const merged = { ...DEFAULTS } as Record<string, unknown>;
    const saved = this.properties as unknown as Record<string, unknown>;
    Object.keys(saved).forEach((k) => {
      if (saved[k] !== undefined && saved[k] !== null) merged[k] = saved[k];
    });
    const result = merged as unknown as IPageRollupWebPartProps;
    result.listId = this.properties.listId || this._autoListId;
    return result;
  }

  /** Before the editor changes anything, fill in the defaults so the property pane shows the real values. */
  private _materialiseDefaults(): void {
    const effective = this._settings();
    const saved = this.properties as unknown as Record<string, unknown>;
    Object.keys(DEFAULTS).forEach((k) => {
      if (saved[k] === undefined || saved[k] === null) saved[k] = (effective as unknown as Record<string, unknown>)[k];
    });
    if (!this.properties.listId && this._autoListId) {
      this.properties.listId = this._autoListId;
    }
  }

  public render(): void {
    const p = this._settings();
    // The service works this out from the library's columns, so the query stays the same once columns load
    // (a changed query would mean a second fetch).
    const isPagesLibrary = false;

    const overlayFields = [p.overlayField1, p.overlayField2, p.overlayField3, p.overlayField4, p.overlayField5].filter((f) => !!f);
    const fieldTitles: Record<string, string> = {};
    this._fields.forEach((f) => {
      fieldTitles[f.internalName] = f.title;
    });

    const query: IRollupQuery = {
      siteUrl: p.siteUrl,
      listId: p.listId,
      isPagesLibrary,
      publishedOnly: p.publishedOnly,
      newsOnly: p.newsOnly,
      filterMatch: p.filterMatch,
      filters: [
        { field: p.filter1Field, operator: p.filter1Operator, values: parseFilterValues(p.filter1Value) },
        { field: p.filter2Field, operator: p.filter2Operator, values: parseFilterValues(p.filter2Value) },
        { field: p.filter3Field, operator: p.filter3Operator, values: parseFilterValues(p.filter3Value) }
      ].filter((f) => !!f.field),
      dateField: p.dateField,
      dateRange: p.dateRange,
      lastDays: p.lastDays,
      sortBy: p.sortBy,
      maxItems: Math.max(1, p.maxItems || 12),
      displayFields: overlayFields,
      thumbnailField: p.thumbnailField,
      includeViews: p.showViews,
      includeSocial: p.showLikes || p.showComments,
      newDays: p.highlightNew ? Math.max(1, p.newDays || 7) : 0,
      viewsPeriod: p.viewsPeriod,
      detailsField: p.hoverEffect === 'reveal' ? p.detailsField : ''
    };

    const display: IRollupDisplay = {
      layout: p.layout,
      showOverlay: p.showOverlay,
      overlayStyle: p.overlayStyle,
      overlayFields,
      fieldTitles,
      showFieldLabels: p.showFieldLabels,
      showViews: p.showViews,
      viewsPeriod: p.viewsPeriod,
      showLikes: p.showLikes,
      showComments: p.showComments,
      highlightNew: p.highlightNew,
      newLabel: (p.newLabel || '').trim() || strings.NewDefaultLabel,
      newGlow: p.newGlow,
      openInNewTab: p.openInNewTab,
      imageQuality: p.imageQuality,
      imageFit: p.imageFit,
      hoverEffect: p.hoverEffect,
      detailsField: p.hoverEffect === 'reveal' ? p.detailsField : ''
    };

    const element: React.ReactElement<IPageRollupProps> = React.createElement(PageRollup, {
      service: this._service,
      query,
      display,
      size: {
        sizeMode: p.sizeMode,
        fixedHeight: p.fixedHeight,
        columns: p.columns,
        tileSize: p.tileSize,
        aspectRatio: p.aspectRatio,
        gap: p.gap,
        cornerRadius: p.cornerRadius
      },
      hoverEffect: p.hoverEffect,
      carousel: {
        autoplay: p.carouselAutoplay,
        interval: p.carouselInterval,
        speed: p.carouselSpeed,
        pauseOnHover: p.carouselPauseOnHover,
        loop: p.carouselLoop,
        arrows: p.carouselArrows,
        dots: p.carouselDots,
        slideWidth: p.carouselSlideWidth,
        background: p.carouselBackground,
        bgColor: p.carouselBgColor,
        bgColor2: p.carouselBgColor2,
        bgImage: p.carouselBgImage,
        bgImageBlur: p.carouselBgImageBlur,
        bgDim: p.carouselBgDim,
        tickerRows: p.tickerRows
      },
      title: p.title,
      titleSize: p.titleSize,
      titleAlign: p.titleAlign,
      titlePlaceholder: strings.InlineTitlePlaceholder,
      onTitleChange: (value: string) => {
        // Save as you type; no re-render needed because the heading keeps its own text
        this.properties.title = value;
      },
      onTitleFormatChange: (size, align) => {
        this.properties.titleSize = size;
        this.properties.titleAlign = align;
        this.render();
        this.context.propertyPane.refresh();
      },
      dateLabel: describeDateWindow(p.dateRange, p.lastDays),
      isConfigured: !!p.listId,
      pending: !p.listId && this._booting,
      isEditMode: this.displayMode === DisplayMode.Edit,
      onConfigure: () => this.context.propertyPane.open(),
      themeVars: this._themeVars()
    });

    ReactDom.render(element, this.domElement);
  }

  private _themeVars(): Record<string, string> {
    const t = this._theme;
    const s = (t && t.semanticColors) || {};
    const pal = (t && t.palette) || {};
    return {
      '--pr-primary': pal.themePrimary || '#0f6cbd',
      '--pr-primary-soft': pal.themeLighter || 'rgba(15,108,189,0.12)',
      '--pr-on-primary': pal.white || '#ffffff',
      '--pr-text': s.bodyText || '#242424',
      '--pr-subtle': s.bodySubtext || '#616161',
      '--pr-card': s.bodyBackground || '#ffffff',
      '--pr-border': s.bodyDivider || s.variantBorder || '#e0e0e0'
    };
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // ----- Property pane -------------------------------------------------------

  protected onPropertyPaneConfigurationStart(): void {
    this._materialiseDefaults();
    this.context.propertyPane.refresh();
    if (this._lists.length === 0) {
      this._loadLists().catch(() => undefined);
    }
    if (this.properties.listId && this._fields.length === 0) {
      this._loadFields().catch(() => undefined);
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (oldValue === newValue) {
      return;
    }
    if (propertyPath === 'siteUrl') {
      this.properties.listId = '';
      this._lists = [];
      this._fields = [];
      this._loadLists().catch(() => undefined);
    }
    if (propertyPath === 'listId') {
      this._fields = [];
      this._loadFields(true).catch(() => undefined);
    }
  }

  /** Site Pages normally; the Documents library when added from the "Document Rollup" toolbox entry. */
  private _defaultLibrary(lists: IRollupList[]): IRollupList | undefined {
    if (this.properties.startWith === 'documents') {
      const docs = lists.filter((l) => l.baseTemplate === 101);
      return docs.filter((l) => /^(documents|shared documents)$/i.test(l.title))[0] || docs[0];
    }
    return lists.filter((l) => l.baseTemplate === SITE_PAGES_TEMPLATE)[0];
  }

  private async _loadLists(): Promise<void> {
    this._loadingLists = true;
    this._sourceError = '';
    this.context.propertyPane.refresh();
    try {
      this._lists = await this._service.getLists(this.properties.siteUrl);
      if (!this.properties.listId) {
        const pages = this._defaultLibrary(this._lists);
        if (pages) {
          this.properties.listId = pages.id;
          await this._loadFields(true);
        }
      }
    } catch (e) {
      this._lists = [];
      this._sourceError = `${strings.SiteLoadError} ${(e as Error).message}`;
    }
    this._loadingLists = false;
    this.context.propertyPane.refresh();
    this.render();
  }

  private async _loadFields(applyDefaults?: boolean): Promise<void> {
    if (!this.properties.listId) {
      return;
    }
    this._loadingFields = true;
    this.context.propertyPane.refresh();
    try {
      this._fields = await this._service.getFields(this.properties.siteUrl, this.properties.listId);
      const has = (n: string): boolean => this._fields.some((f) => f.internalName === n);
      if (applyDefaults || !has(this.properties.dateField)) {
        this.properties.dateField = has('FirstPublishedDate') ? 'FirstPublishedDate' : 'Created';
      }
      if (applyDefaults) {
        this.properties.overlayField2 = has('Description') ? 'Description' : has('Modified') ? 'Modified' : '';
        this.properties.thumbnailField = '';
        this.properties.detailsField = has('Description') ? 'Description' : '';
        ['filter1Field', 'filter2Field', 'filter3Field'].forEach((k) => {
          const key = k as 'filter1Field' | 'filter2Field' | 'filter3Field';
          if (this.properties[key] && !has(this.properties[key])) {
            this.properties[key] = '';
          }
        });
      }
    } catch {
      this._fields = [];
    }
    this._loadingFields = false;
    this.context.propertyPane.refresh();
    this.render();
  }

  private _fieldOptions(filter?: (f: IRollupField) => boolean, noneLabel: string = strings.NoneOption): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [{ key: NONE, text: noneLabel }];
    this._fields.filter(filter || (() => true)).forEach((f) => options.push({ key: f.internalName, text: `${f.title}` }));
    return options;
  }

  private _filterFields(n: 1 | 2 | 3): IPropertyPaneField<unknown>[] {
    const fieldKey = `filter${n}Field` as 'filter1Field';
    const opKey = `filter${n}Operator` as 'filter1Operator';
    const out: IPropertyPaneField<unknown>[] = [
      PropertyPaneDropdown(fieldKey, {
        label: `${strings.FilterLabel} ${n}`,
        options: this._fieldOptions((f) => f.type !== 'Thumbnail' && f.type !== 'URL'),
        disabled: this._loadingFields
      })
    ];
    if (this.properties[fieldKey]) {
      const op: FilterOperator = this.properties[opKey];
      out.push(
        PropertyPaneDropdown(opKey, {
          label: strings.OperatorLabel,
          options: [
            { key: 'eq', text: strings.OpEquals },
            { key: 'neq', text: strings.OpNotEquals },
            { key: 'contains', text: strings.OpContains },
            { key: 'empty', text: strings.OpEmpty },
            { key: 'notEmpty', text: strings.OpNotEmpty }
          ]
        })
      );
      if (op !== 'empty' && op !== 'notEmpty') {
        out.push(
          PropertyPaneTextField(`filter${n}Value`, {
            label: strings.ValueLabel,
            placeholder: strings.ValuePlaceholder,
            description: n === 1 ? strings.ValueDescription : undefined,
            deferredValidationTime: 500
          })
        );
      }
    }
    return out;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const p = this._settings();
    const list = this._lists.filter((l) => l.id === p.listId)[0];
    const isPages = list ? list.baseTemplate === SITE_PAGES_TEMPLATE : this._fields.some((f) => f.internalName === 'BannerImageUrl');

    const listOptions: IPropertyPaneDropdownOption[] = this._lists.map((l) => ({
      key: l.id,
      text: l.baseTemplate === SITE_PAGES_TEMPLATE ? `${l.title} (${strings.PagesTag})` : l.title
    }));

    const sourceFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneTextField('title', { label: strings.TitleLabel, placeholder: strings.TitlePlaceholder, description: strings.TitleInlineHelp }),
      PropertyPaneDropdown('titleSize', {
        label: strings.TitleSizeLabel,
        options: [
          { key: 'small', text: strings.SizeSmall },
          { key: 'medium', text: strings.SizeMedium },
          { key: 'large', text: strings.SizeLarge },
          { key: 'xlarge', text: strings.SizeXLarge }
        ]
      }),
      PropertyPaneChoiceGroup('titleAlign', {
        label: strings.TitleAlignLabel,
        options: [
          { key: 'left', text: strings.AlignLeft, iconProps: { officeFabricIconFontName: 'AlignLeft' } },
          { key: 'center', text: strings.AlignCenter, iconProps: { officeFabricIconFontName: 'AlignCenter' } },
          { key: 'right', text: strings.AlignRight, iconProps: { officeFabricIconFontName: 'AlignRight' } }
        ]
      }),
      PropertyPaneTextField('siteUrl', {
        label: strings.SiteUrlLabel,
        placeholder: this.context.pageContext.web.absoluteUrl,
        description: strings.SiteUrlDescription,
        deferredValidationTime: 800,
        onGetErrorMessage: (v: string) => (!v || /^https:\/\/[^/]+/i.test(v.trim()) ? '' : strings.SiteUrlError)
      }),
      PropertyPaneDropdown('listId', {
        label: strings.LibraryLabel,
        options: listOptions,
        disabled: this._loadingLists || listOptions.length === 0,
        selectedKey: p.listId
      })
    ];
    if (this._loadingLists) {
      sourceFields.push(PropertyPaneLabel('loadingLists', { text: strings.Loading }));
    }
    if (this._sourceError) {
      sourceFields.push(PropertyPaneLabel('sourceError', { text: this._sourceError }));
    }
    if (isPages) {
      sourceFields.push(
        PropertyPaneToggle('publishedOnly', { label: strings.PublishedOnlyLabel, onText: strings.On, offText: strings.Off }),
        PropertyPaneToggle('newsOnly', { label: strings.NewsOnlyLabel, onText: strings.On, offText: strings.Off })
      );
    }

    const filterFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneLabel('filterHelp', { text: strings.FilterHelp }),
      ...this._filterFields(1),
      ...(p.filter1Field ? this._filterFields(2) : []),
      ...(p.filter1Field && p.filter2Field ? this._filterFields(3) : [])
    ];
    if (p.filter1Field && p.filter2Field) {
      filterFields.push(
        PropertyPaneChoiceGroup('filterMatch', {
          label: strings.FilterMatchLabel,
          options: [
            { key: 'all', text: strings.MatchAll },
            { key: 'any', text: strings.MatchAny }
          ]
        })
      );
    }

    const dateFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneDropdown('dateField', {
        label: strings.DateFieldLabel,
        options: this._fieldOptions((f) => f.type === 'DateTime').filter((o) => o.key !== NONE),
        selectedKey: p.dateField
      }),
      PropertyPaneDropdown('dateRange', {
        label: strings.DateRangeLabel,
        options: [
          { key: 'all', text: strings.RangeAll },
          { key: 'lastDays', text: strings.RangeLastDays },
          { key: 'thisQuarter', text: strings.RangeThisQuarter },
          { key: 'lastQuarter', text: strings.RangeLastQuarter },
          { key: 'thisYear', text: strings.RangeThisYear }
        ]
      })
    ];
    if (p.dateRange === 'lastDays') {
      dateFields.push(PropertyPaneSlider('lastDays', { label: strings.LastDaysLabel, min: 1, max: 365, step: 1, showValue: true }));
    }
    dateFields.push(
      PropertyPaneDropdown('sortBy', {
        label: strings.SortLabel,
        options: [
          { key: 'dateDesc', text: strings.SortNewest },
          { key: 'dateAsc', text: strings.SortOldest },
          { key: 'title', text: strings.SortTitle },
          { key: 'views', text: strings.SortViews },
          { key: 'likes', text: strings.SortLikes }
        ]
      }),
      PropertyPaneSlider('maxItems', { label: strings.MaxItemsLabel, min: 1, max: 50, step: 1, showValue: true })
    );

    const overlayFieldOptions = this._fieldOptions((f) => f.type !== 'Thumbnail');
    const displayFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneChoiceGroup('layout', {
        label: strings.LayoutLabel,
        options: [
          { key: 'tiles', text: strings.LayoutTiles, iconProps: { officeFabricIconFontName: 'Tiles' } },
          { key: 'cards', text: strings.LayoutCards, iconProps: { officeFabricIconFontName: 'GridViewMedium' } },
          { key: 'hero', text: strings.LayoutHero, iconProps: { officeFabricIconFontName: 'ViewDashboard' } },
          { key: 'mosaic', text: strings.LayoutMosaic, iconProps: { officeFabricIconFontName: 'GridViewLarge' } },
          { key: 'filmstrip', text: strings.LayoutFilmstrip, iconProps: { officeFabricIconFontName: 'TripleColumn' } },
          { key: 'list', text: strings.LayoutList, iconProps: { officeFabricIconFontName: 'ViewList' } },
          { key: 'spotlight', text: strings.LayoutSpotlight, iconProps: { officeFabricIconFontName: 'Slideshow' } },
          { key: 'social', text: strings.LayoutSocial, iconProps: { officeFabricIconFontName: 'PhotoCollection' } },
          { key: 'ticker', text: strings.LayoutTicker, iconProps: { officeFabricIconFontName: 'SortUp' } }
        ]
      }),
      PropertyPaneLabel('layoutHelp', { text: (strings as unknown as Record<string, string>)[`LayoutHelp_${p.layout}`] || '' }),
      PropertyPaneDropdown('hoverEffect', {
        label: strings.HoverLabel,
        options: [
          { key: 'zoom', text: strings.HoverZoom },
          { key: 'lift', text: strings.HoverLift },
          { key: 'reveal', text: strings.HoverReveal },
          { key: 'none', text: strings.HoverNone }
        ]
      }),
      ...(p.hoverEffect === 'reveal'
        ? [
            PropertyPaneDropdown('detailsField', {
              label: strings.DetailsFieldLabel,
              options: this._fieldOptions((f) => f.type !== 'Thumbnail', strings.DetailsNone),
              selectedKey: p.detailsField
            }),
            PropertyPaneLabel('detailsHelp', { text: strings.DetailsHelp })
          ]
        : []),
      PropertyPaneDropdown('thumbnailField', {
        label: strings.ThumbnailLabel,
        options: this._fieldOptions((f) => f.type === 'Thumbnail' || f.type === 'URL', strings.ThumbnailAuto)
      }),
      PropertyPaneDropdown('imageFit', {
        label: strings.ImageFitLabel,
        options: [
          { key: 'cover', text: strings.FitCover },
          { key: 'contain', text: strings.FitContain }
        ]
      }),
      PropertyPaneDropdown('imageQuality', {
        label: strings.ImageQualityLabel,
        options: [
          { key: 'sharp', text: strings.ImageSharp },
          { key: 'fast', text: strings.ImageFast },
          { key: 'original', text: strings.ImageOriginal }
        ]
      }),
      PropertyPaneHorizontalRule(),
      PropertyPaneToggle('showOverlay', { label: strings.ShowOverlayLabel, onText: strings.On, offText: strings.Off })
    ];
    if (p.showOverlay) {
      if (p.layout !== 'cards' && p.layout !== 'list') {
        displayFields.push(
          PropertyPaneDropdown('overlayStyle', {
            label: strings.OverlayStyleLabel,
            options: [
              { key: 'gradient', text: strings.OverlayGradient },
              { key: 'glass', text: strings.OverlayGlass },
              { key: 'solid', text: strings.OverlaySolid }
            ]
          })
        );
      }
      displayFields.push(
        PropertyPaneDropdown('overlayField1', { label: strings.OverlayField1Label, options: overlayFieldOptions, selectedKey: p.overlayField1 }),
        PropertyPaneDropdown('overlayField2', { label: `${strings.OverlayFieldLabel} 2`, options: overlayFieldOptions, selectedKey: p.overlayField2 }),
        PropertyPaneDropdown('overlayField3', { label: `${strings.OverlayFieldLabel} 3`, options: overlayFieldOptions, selectedKey: p.overlayField3 }),
        PropertyPaneDropdown('overlayField4', { label: `${strings.OverlayFieldLabel} 4`, options: overlayFieldOptions, selectedKey: p.overlayField4 }),
        PropertyPaneDropdown('overlayField5', { label: `${strings.OverlayFieldLabel} 5`, options: overlayFieldOptions, selectedKey: p.overlayField5 }),
        PropertyPaneToggle('showFieldLabels', { label: strings.ShowFieldLabels, onText: strings.On, offText: strings.Off })
      );
    }
    displayFields.push(
      PropertyPaneHorizontalRule(),
      PropertyPaneToggle('showViews', { label: strings.ShowViewsLabel, onText: strings.On, offText: strings.Off })
    );
    if (p.showViews) {
      displayFields.push(
        PropertyPaneDropdown('viewsPeriod', {
          label: strings.ViewsPeriodLabel,
          options: [
            { key: 'lifetime', text: strings.ViewsLifetime },
            { key: 'recent', text: strings.ViewsRecent }
          ]
        })
      );
    }
    displayFields.push(
      PropertyPaneToggle('showLikes', { label: strings.ShowLikesLabel, onText: strings.On, offText: strings.Off }),
      PropertyPaneToggle('showComments', { label: strings.ShowCommentsLabel, onText: strings.On, offText: strings.Off })
    );
    if (p.showLikes || p.showComments) {
      displayFields.push(PropertyPaneLabel('socialHelp', { text: strings.SocialHelp }));
    }
    displayFields.push(
      PropertyPaneHorizontalRule(),
      PropertyPaneToggle('highlightNew', { label: strings.HighlightNewLabel, onText: strings.On, offText: strings.Off })
    );
    if (p.highlightNew) {
      const dateTitle = (this._fields.filter((f) => f.internalName === p.dateField)[0] || { title: p.dateField || 'Created' }).title;
      displayFields.push(
        PropertyPaneSlider('newDays', { label: strings.NewDaysLabel, min: 1, max: 90, step: 1, showValue: true }),
        PropertyPaneLabel('newHelp', { text: strings.NewHelp.replace('{0}', dateTitle) }),
        PropertyPaneTextField('newLabel', { label: strings.NewLabelLabel, placeholder: strings.NewDefaultLabel, maxLength: 16 }),
        PropertyPaneToggle('newGlow', { label: strings.NewGlowLabel, onText: strings.On, offText: strings.Off })
      );
    }
    displayFields.push(PropertyPaneToggle('openInNewTab', { label: strings.NewTabLabel, onText: strings.On, offText: strings.Off }));

    const isCarousel = p.layout === 'spotlight' || p.layout === 'social' || p.layout === 'ticker';
    const carouselFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneToggle('carouselAutoplay', { label: strings.AutoplayLabel, onText: strings.On, offText: strings.Off })
    ];
    if (p.carouselAutoplay) {
      carouselFields.push(
        PropertyPaneSlider('carouselInterval', { label: strings.IntervalLabel, min: 2, max: 20, step: 1, showValue: true }),
        PropertyPaneToggle('carouselPauseOnHover', { label: strings.PauseLabel, onText: strings.On, offText: strings.Off })
      );
    }
    carouselFields.push(
      PropertyPaneSlider('carouselSpeed', { label: strings.SpeedLabel, min: 200, max: 2000, step: 50, showValue: true }),
      PropertyPaneToggle('carouselLoop', { label: strings.LoopLabel, onText: strings.On, offText: strings.Off }),
      PropertyPaneToggle('carouselArrows', { label: strings.ArrowsLabel, onText: strings.On, offText: strings.Off })
    );
    if (p.layout !== 'ticker') {
      carouselFields.push(PropertyPaneToggle('carouselDots', { label: strings.DotsLabel, onText: strings.On, offText: strings.Off }));
    }
    if (p.layout === 'ticker') {
      carouselFields.push(PropertyPaneSlider('tickerRows', { label: strings.TickerRowsLabel, min: 1, max: 6, step: 1, showValue: true }));
    }
    if (p.layout === 'spotlight') {
      carouselFields.push(
        PropertyPaneSlider('carouselSlideWidth', { label: strings.SlideWidthLabel, min: 35, max: 90, step: 5, showValue: true }),
        PropertyPaneDropdown('carouselBackground', {
          label: strings.BackgroundLabel,
          options: [
            { key: 'blur', text: strings.BgBlur },
            { key: 'image', text: strings.BgImage },
            { key: 'gradient', text: strings.BgGradient },
            { key: 'color', text: strings.BgColor },
            { key: 'theme', text: strings.BgTheme },
            { key: 'none', text: strings.BgNone }
          ]
        })
      );
      const colourCheck = (v: string): string => (!v || /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(v.trim()) ? '' : strings.ColourError);
      if (p.carouselBackground === 'color' || p.carouselBackground === 'gradient') {
        carouselFields.push(
          PropertyPaneTextField('carouselBgColor', {
            label: p.carouselBackground === 'gradient' ? strings.BgColourFrom : strings.BgColourLabel,
            placeholder: '#1b1f3b',
            onGetErrorMessage: colourCheck,
            deferredValidationTime: 400
          })
        );
      }
      if (p.carouselBackground === 'image') {
        carouselFields.push(
          PropertyPaneTextField('carouselBgImage', {
            label: strings.BgImageLabel,
            placeholder: 'https://nats.sharepoint.com/sites/…/SiteAssets/background.jpg',
            description: strings.BgImageHelp,
            multiline: false,
            deferredValidationTime: 600,
            onGetErrorMessage: (v: string) => (!v || /^(https:\/\/|\/)/i.test(v.trim()) ? '' : strings.BgImageError)
          }),
          PropertyPaneToggle('carouselBgImageBlur', { label: strings.BgImageBlurLabel, onText: strings.On, offText: strings.Off })
        );
      }
      if (p.carouselBackground === 'image' || p.carouselBackground === 'blur') {
        carouselFields.push(PropertyPaneSlider('carouselBgDim', { label: strings.BgDimLabel, min: 0, max: 80, step: 5, showValue: true }));
      }
      if (p.carouselBackground === 'gradient') {
        carouselFields.push(
          PropertyPaneTextField('carouselBgColor2', { label: strings.BgColourTo, placeholder: '#5b2a86', onGetErrorMessage: colourCheck, deferredValidationTime: 400 })
        );
      }
    }

    const sizeFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneChoiceGroup('sizeMode', {
        label: strings.SizeModeLabel,
        options: [
          { key: 'auto', text: strings.SizeAuto },
          { key: 'fixed', text: strings.SizeFixed }
        ]
      })
    ];
    if (p.sizeMode === 'fixed') {
      sizeFields.push(PropertyPaneSlider('fixedHeight', { label: strings.FixedHeightLabel, min: 150, max: 1400, step: 10, showValue: true }));
    }
    sizeFields.push(
      PropertyPaneSlider('columns', { label: strings.ColumnsLabel, min: 0, max: 6, step: 1, showValue: true }),
      PropertyPaneLabel('columnsHelp', { text: strings.ColumnsHelp })
    );
    if (!p.columns) {
      sizeFields.push(PropertyPaneSlider('tileSize', { label: strings.TileSizeLabel, min: 140, max: 520, step: 10, showValue: true }));
    }
    sizeFields.push(
      PropertyPaneDropdown('aspectRatio', {
        label: strings.AspectLabel,
        options: [
          { key: '16:9', text: '16:9 (widescreen)' },
          { key: '4:3', text: '4:3' },
          { key: '1:1', text: '1:1 (square)' },
          { key: '3:4', text: '3:4 (portrait)' }
        ]
      }),
      PropertyPaneSlider('gap', { label: strings.GapLabel, min: 0, max: 40, step: 2, showValue: true }),
      PropertyPaneSlider('cornerRadius', { label: strings.RadiusLabel, min: 0, max: 28, step: 1, showValue: true })
    );

    return {
      pages: [
        {
          header: { description: strings.PageSource },
          displayGroupsAsAccordion: true,
          groups: [
            { groupName: strings.GroupSource, groupFields: sourceFields },
            { groupName: strings.GroupFilters, groupFields: filterFields, isCollapsed: !p.filter1Field },
            { groupName: strings.GroupDates, groupFields: dateFields }
          ]
        },
        {
          header: { description: strings.PageDisplay },
          groups: [
            { groupName: strings.GroupDisplay, groupFields: displayFields },
            ...(isCarousel ? [{ groupName: strings.GroupCarousel, groupFields: carouselFields }] : [])
          ]
        },
        {
          header: { description: strings.PageSize },
          groups: [{ groupName: strings.GroupSize, groupFields: sizeFields }]
        }
      ]
    };
  }
}
