import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IReadonlyTheme,
  ThemeChangedEventArgs,
  ThemeProvider
} from '@microsoft/sp-component-base';

import * as strings from 'LikeWebPartStrings';
import { LikeButton } from './components/LikeButton';
import { ILikeButtonProps } from './components/ILikeButtonProps';
import { ILikeService } from './services/ILikeService';
import { PageLikeService } from './services/PageLikeService';
import { ListLikeService } from './services/ListLikeService';
import { ILikeWebPartProps } from './models';

const HEART_RED = '#d13438';

export default class LikeWebPart extends BaseClientSideWebPart<ILikeWebPartProps> {
  private _themeProvider: ThemeProvider | undefined;
  private _themeVariant: IReadonlyTheme | undefined;
  private _service: ILikeService | undefined;
  private _serviceKey: string | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._themeVariant = this._themeProvider.tryGetTheme();
    this._themeProvider.themeChangedEvent.add(this, this._onThemeChanged);
  }

  public render(): void {
    const p = this.properties;
    const { service, unavailableMessage } = this._getService();

    const element = React.createElement<ILikeButtonProps>(LikeButton, {
      service,
      unavailableMessage,
      currentUserName: this.context.pageContext.user.displayName,
      listTitle: p.listTitle || 'PageLikes',
      displayText: p.displayText,
      likedText: p.likedText,
      iconType: p.iconType || 'thumb',
      customEmoji: p.customEmoji,
      size: p.size || 'medium',
      alignment: p.alignment || 'left',
      layout: p.layout || 'inline',
      buttonStyle: p.buttonStyle || 'outline',
      accentColor: this._resolveAccent(),
      themeVars: this._getThemeVars(),
      showCount: p.showCount !== false,
      countLabelSingular: p.countLabelSingular ?? '',
      countLabelPlural: p.countLabelPlural ?? '',
      showLikers: !!p.showLikers,
      allowUnlike: p.allowUnlike !== false,
      containerStyle: p.containerStyle || 'none',
      backgroundColor: LikeWebPart._validColorOrUndefined(p.backgroundColor),
      textColor: LikeWebPart._validColorOrUndefined(p.textColor),
      cornerRadius: p.cornerRadius || 'medium',
      containerWidth: p.containerWidth || 'full',
      isEditMode: this.displayMode === DisplayMode.Edit
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    this._themeProvider?.themeChangedEvent.remove(this, this._onThemeChanged);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _onThemeChanged(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  /** Map the (section-aware) theme onto the CSS variables used by the button. */
  private _getThemeVars(): Record<string, string> | undefined {
    const t = this._themeVariant;
    if (!t) {
      return undefined;
    }
    const vars: Record<string, string> = {};
    if (t.palette?.themePrimary) vars['--like-accent'] = t.palette.themePrimary;
    if (t.semanticColors?.bodyText) vars['--like-text'] = t.semanticColors.bodyText;
    if (t.semanticColors?.bodySubtext) vars['--like-subtle-text'] = t.semanticColors.bodySubtext;
    return vars;
  }

  private _resolveAccent(): string | undefined {
    const custom = (this.properties.accentColor || '').trim();
    if (custom && LikeWebPart._isValidColor(custom)) {
      return custom;
    }
    return this.properties.iconType === 'heart' ? HEART_RED : undefined;
  }

  private static _validColorOrUndefined(value: string | undefined): string | undefined {
    const v = (value || '').trim();
    return v && LikeWebPart._isValidColor(v) ? v : undefined;
  }

  private static _colorError(value: string): string {
    return !value || !value.trim() || LikeWebPart._isValidColor(value.trim()) ? '' : strings.InvalidColor;
  }

  private static _isValidColor(value: string): boolean {
    return typeof CSS !== 'undefined' && typeof CSS.supports === 'function'
      ? CSS.supports('color', value)
      : /^#([0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(value);
  }

  /** Default like key for list storage: page URL + this web part instance. */
  private _defaultLikeKey(): string {
    const path = (this.context.pageContext.legacyPageContext?.serverRequestPath as string) ||
      window.location.pathname;
    return `${path.toLowerCase()}#${this.instanceId}`;
  }

  /** Create the storage service, re-using it while its settings are unchanged. */
  private _getService(): { service?: ILikeService; unavailableMessage?: string } {
    const ctx = this.context.pageContext;
    const p = this.properties;
    const webUrl = ctx.web.absoluteUrl;

    if (p.storage === 'list') {
      const listTitle = (p.listTitle || 'PageLikes').trim();
      const likeKey = (p.likeKey || '').trim() || this._defaultLikeKey();
      const userId = Number(ctx.legacyPageContext?.userId) || 0;
      const key = ['list', webUrl, listTitle, likeKey, userId].join('|');
      if (key !== this._serviceKey) {
        this._service = new ListLikeService(this.context.spHttpClient, webUrl, listTitle, likeKey, userId);
        this._serviceKey = key;
      }
      return { service: this._service };
    }

    const itemId = ctx.listItem?.id;
    if (!itemId) {
      this._service = undefined;
      this._serviceKey = undefined;
      return { unavailableMessage: strings.NoPageContext };
    }
    const key = ['page', webUrl, itemId].join('|');
    if (key !== this._serviceKey) {
      this._service = new PageLikeService(this.context.spHttpClient, webUrl, itemId);
      this._serviceKey = key;
    }
    return { service: this._service };
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const p = this.properties;

    const iconFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneChoiceGroup('iconType', {
        label: strings.IconTypeLabel,
        options: [
          { key: 'thumb', text: strings.IconThumb, iconProps: { officeFabricIconFontName: 'Like' } },
          { key: 'heart', text: strings.IconHeart, iconProps: { officeFabricIconFontName: 'Heart' } },
          { key: 'star', text: strings.IconStar, iconProps: { officeFabricIconFontName: 'FavoriteStar' } },
          { key: 'emoji', text: strings.IconEmoji, iconProps: { officeFabricIconFontName: 'Emoji2' } }
        ]
      })
    ];
    if (p.iconType === 'emoji') {
      iconFields.push(PropertyPaneTextField('customEmoji', { label: strings.CustomEmojiLabel, placeholder: '👏' }));
    }

    const countFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneToggle('showCount', { label: strings.ShowCountLabel, onText: strings.On, offText: strings.Off })
    ];
    if (p.showCount !== false) {
      countFields.push(
        PropertyPaneTextField('countLabelSingular', { label: strings.CountSingularLabel }),
        PropertyPaneTextField('countLabelPlural', { label: strings.CountPluralLabel })
      );
    }
    countFields.push(
      PropertyPaneToggle('showLikers', { label: strings.ShowLikersLabel, onText: strings.On, offText: strings.Off }),
      PropertyPaneToggle('allowUnlike', { label: strings.AllowUnlikeLabel, onText: strings.On, offText: strings.Off })
    );

    const backgroundFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneDropdown('containerStyle', {
        label: strings.ContainerStyleLabel,
        options: [
          { key: 'none', text: strings.ContainerNone },
          { key: 'soft', text: strings.ContainerSoft },
          { key: 'card', text: strings.ContainerCard },
          { key: 'outline', text: strings.ContainerOutline },
          { key: 'frosted', text: strings.ContainerFrosted },
          { key: 'accent', text: strings.ContainerAccent }
        ]
      })
    ];
    if (p.containerStyle && p.containerStyle !== 'none') {
      backgroundFields.push(
        PropertyPaneTextField('backgroundColor', {
          label: strings.BackgroundColorLabel,
          description: strings.BackgroundColorDescription,
          placeholder: '#ffffff',
          onGetErrorMessage: LikeWebPart._colorError
        }),
        PropertyPaneDropdown('cornerRadius', {
          label: strings.CornerRadiusLabel,
          options: [
            { key: 'none', text: strings.RadiusNone },
            { key: 'small', text: strings.RadiusSmall },
            { key: 'medium', text: strings.RadiusMedium },
            { key: 'large', text: strings.RadiusLarge }
          ]
        }),
        PropertyPaneChoiceGroup('containerWidth', {
          label: strings.ContainerWidthLabel,
          options: [
            { key: 'full', text: strings.WidthFull },
            { key: 'fit', text: strings.WidthFit }
          ]
        })
      );
    }
    backgroundFields.push(
      PropertyPaneTextField('textColor', {
        label: strings.TextColorLabel,
        description: strings.TextColorDescription,
        placeholder: '#323130',
        onGetErrorMessage: LikeWebPart._colorError
      })
    );

    const storageFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneChoiceGroup('storage', {
        label: strings.StorageLabel,
        options: [
          { key: 'page', text: strings.StoragePage },
          { key: 'list', text: strings.StorageList }
        ]
      }),
      PropertyPaneLabel('storageHelp', {
        text: p.storage === 'list' ? strings.StorageListDescription : strings.StoragePageDescription
      })
    ];
    if (p.storage === 'list') {
      storageFields.push(
        PropertyPaneTextField('listTitle', { label: strings.ListTitleLabel, placeholder: 'PageLikes' }),
        PropertyPaneTextField('likeKey', {
          label: strings.LikeKeyLabel,
          description: strings.LikeKeyDescription,
          placeholder: this._defaultLikeKey()
        })
      );
    }

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.TextGroupName,
              groupFields: [
                PropertyPaneTextField('displayText', { label: strings.DisplayTextLabel, multiline: true, rows: 2 }),
                PropertyPaneTextField('likedText', {
                  label: strings.LikedTextLabel,
                  description: strings.LikedTextDescription,
                  multiline: true,
                  rows: 2
                })
              ]
            },
            { groupName: strings.IconGroupName, groupFields: iconFields },
            {
              groupName: strings.AppearanceGroupName,
              groupFields: [
                PropertyPaneDropdown('size', {
                  label: strings.SizeLabel,
                  options: [
                    { key: 'small', text: strings.SizeSmall },
                    { key: 'medium', text: strings.SizeMedium },
                    { key: 'large', text: strings.SizeLarge },
                    { key: 'xlarge', text: strings.SizeXLarge }
                  ]
                }),
                PropertyPaneChoiceGroup('alignment', {
                  label: strings.AlignmentLabel,
                  options: [
                    { key: 'left', text: strings.AlignLeft, iconProps: { officeFabricIconFontName: 'AlignLeft' } },
                    { key: 'center', text: strings.AlignCenter, iconProps: { officeFabricIconFontName: 'AlignCenter' } },
                    { key: 'right', text: strings.AlignRight, iconProps: { officeFabricIconFontName: 'AlignRight' } }
                  ]
                }),
                PropertyPaneDropdown('layout', {
                  label: strings.LayoutLabel,
                  options: [
                    { key: 'inline', text: strings.LayoutInline },
                    { key: 'stacked', text: strings.LayoutStacked }
                  ]
                }),
                PropertyPaneDropdown('buttonStyle', {
                  label: strings.ButtonStyleLabel,
                  options: [
                    { key: 'filled', text: strings.StyleFilled },
                    { key: 'outline', text: strings.StyleOutline },
                    { key: 'subtle', text: strings.StyleSubtle }
                  ]
                }),
                PropertyPaneTextField('accentColor', {
                  label: strings.AccentColorLabel,
                  description: strings.AccentColorDescription,
                  placeholder: p.iconType === 'heart' ? HEART_RED : '#0078d4',
                  onGetErrorMessage: LikeWebPart._colorError
                })
              ]
            },
            { groupName: strings.BackgroundGroupName, groupFields: backgroundFields },
            { groupName: strings.CountGroupName, groupFields: countFields }
          ]
        },
        {
          header: { description: strings.DataPageName },
          groups: [{ groupName: strings.StorageGroupName, groupFields: storageFields }]
        }
      ]
    };
  }
}
