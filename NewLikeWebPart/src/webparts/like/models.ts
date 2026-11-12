export type LikeIconType = 'thumb' | 'heart' | 'star' | 'emoji';
export type LikeSize = 'small' | 'medium' | 'large' | 'xlarge';
export type LikeAlignment = 'left' | 'center' | 'right';
export type LikeLayout = 'inline' | 'stacked';
export type LikeButtonStyle = 'filled' | 'outline' | 'subtle';
export type LikeStorage = 'page' | 'list';

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
  showCount: boolean;
  countLabelSingular: string;
  countLabelPlural: string;
  showLikers: boolean;
  allowUnlike: boolean;
  storage: LikeStorage;
  listTitle: string;
  likeKey: string;
}
