export type LikeIconType = 'thumb' | 'heart' | 'star' | 'emoji';
export type LikeSize = 'small' | 'medium' | 'large' | 'xlarge';
export type LikeAlignment = 'left' | 'center' | 'right';
export type LikeLayout = 'inline' | 'stacked';
export type LikeButtonStyle = 'filled' | 'outline' | 'subtle';
export type LikeStorage = 'page' | 'list';
/** Background behind the whole web part (text + button). */
export type LikeContainerStyle = 'none' | 'soft' | 'card' | 'outline' | 'accent' | 'frosted';
export type LikeCornerRadius = 'none' | 'small' | 'medium' | 'large';
export type LikeContainerWidth = 'full' | 'fit';

export interface ILikeWebPartProps {
  displayText: string;
  likedText: string;
  iconType: LikeIconType;
  customEmoji: string;
  size: LikeSize;
  alignment: LikeAlignment;
  layout: LikeLayout;
  buttonStyle: LikeButtonStyle;
  accentColor: string;
  containerStyle: LikeContainerStyle;
  backgroundColor: string;
  textColor: string;
  cornerRadius: LikeCornerRadius;
  containerWidth: LikeContainerWidth;
  showCount: boolean;
  countLabelSingular: string;
  countLabelPlural: string;
  showLikers: boolean;
  allowUnlike: boolean;
  storage: LikeStorage;
  listTitle: string;
  likeKey: string;
}
