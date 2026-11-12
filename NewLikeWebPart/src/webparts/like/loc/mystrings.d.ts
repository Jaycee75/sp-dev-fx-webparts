declare interface ILikeWebPartStrings {
  PropertyPaneDescription: string;
  DisplayPageName: string;
  DataPageName: string;
  TextGroupName: string;
  IconGroupName: string;
  AppearanceGroupName: string;
  CountGroupName: string;
  StorageGroupName: string;
  DisplayTextLabel: string;
  LikedTextLabel: string;
  LikedTextDescription: string;
  IconTypeLabel: string;
  IconThumb: string;
  IconHeart: string;
  IconStar: string;
  IconEmoji: string;
  CustomEmojiLabel: string;
  SizeLabel: string;
  SizeSmall: string;
  SizeMedium: string;
  SizeLarge: string;
  SizeXLarge: string;
  AlignmentLabel: string;
  AlignLeft: string;
  AlignCenter: string;
  AlignRight: string;
  LayoutLabel: string;
  LayoutInline: string;
  LayoutStacked: string;
  ButtonStyleLabel: string;
  StyleFilled: string;
  StyleOutline: string;
  StyleSubtle: string;
  AccentColorLabel: string;
  AccentColorDescription: string;
  InvalidColor: string;
  ShowCountLabel: string;
  CountSingularLabel: string;
  CountPluralLabel: string;
  ShowLikersLabel: string;
  AllowUnlikeLabel: string;
  StorageLabel: string;
  StoragePage: string;
  StorageList: string;
  StoragePageDescription: string;
  StorageListDescription: string;
  ListTitleLabel: string;
  LikeKeyLabel: string;
  LikeKeyDescription: string;
  On: string;
  Off: string;

  LikeAriaLabel: string;
  UnlikeAriaLabel: string;
  LikedByPrefix: string;
  AndOthers: string;
  LoadError: string;
  SaveError: string;
  AccessDenied: string;
  NoPageContext: string;
  Loading: string;
}

declare module 'LikeWebPartStrings' {
  const strings: ILikeWebPartStrings;
  export = strings;
}
