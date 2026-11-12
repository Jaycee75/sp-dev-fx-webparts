import { ILikeService } from '../services/ILikeService';
import {
  LikeAlignment,
  LikeContainerStyle,
  LikeContainerWidth,
  LikeCornerRadius,
  LikeButtonStyle,
  LikeIconType,
  LikeLayout,
  LikeSize
} from '../models';

export interface ILikeButtonProps {
  /** Undefined when likes can't be stored in the current context. */
  service: ILikeService | undefined;
  /** Message shown when there is no service. */
  unavailableMessage?: string;
  currentUserName: string;
  listTitle: string;

  displayText: string;
  likedText: string;
  iconType: LikeIconType;
  customEmoji: string;
  size: LikeSize;
  alignment: LikeAlignment;
  layout: LikeLayout;
  buttonStyle: LikeButtonStyle;
  /** CSS custom properties derived from the (section) theme. */
  themeVars?: Record<string, string>;
  /** Resolved accent colour; undefined means "use the theme". */
  accentColor: string | undefined;
  showCount: boolean;
  countLabelSingular: string;
  countLabelPlural: string;
  showLikers: boolean;
  allowUnlike: boolean;

  containerStyle: LikeContainerStyle;
  /** Validated CSS colour or undefined. */
  backgroundColor: string | undefined;
  /** Validated CSS colour or undefined. */
  textColor: string | undefined;
  cornerRadius: LikeCornerRadius;
  containerWidth: LikeContainerWidth;
  /** Shows technical error details to page editors. */
  isEditMode: boolean;
}
