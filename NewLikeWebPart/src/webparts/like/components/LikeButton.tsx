import * as React from 'react';
import * as strings from 'LikeWebPartStrings';
import styles from './LikeButton.module.scss';
import { ILikeButtonProps } from './ILikeButtonProps';
import { LikeIcon } from './LikeIcon';
import { ILikeState, LikeServiceError } from '../services/ILikeService';

const MAX_LIKERS_IN_TOOLTIP = 10;
const cls = styles as unknown as Record<string, string | undefined>;

function cx(...classes: (string | false | undefined)[]): string {
  return classes.filter(Boolean).join(' ');
}

/** 999 -> "999", 1234 -> "1.2K", 1500000 -> "1.5M" */
export function formatCount(n: number): string {
  if (n < 1000) {
    return String(n);
  }
  const units: [number, string][] = [[1e6, 'M'], [1e3, 'K']];
  for (const [value, suffix] of units) {
    if (n >= value) {
      const v = n / value;
      return `${v >= 100 ? Math.round(v) : Math.round(v * 10) / 10}${suffix}`;
    }
  }
  return String(n);
}

function format(template: string, ...args: (string | number)[]): string {
  return template.replace(/\{(\d+)\}/g, (m, i) => (args[+i] !== undefined ? String(args[+i]) : m));
}

function likersTooltip(likers: string[], count: number): string | undefined {
  if (!likers.length) {
    return undefined;
  }
  const shown = likers.slice(0, MAX_LIKERS_IN_TOOLTIP);
  const others = Math.max(count, likers.length) - shown.length;
  return `${strings.LikedByPrefix} ${shown.join(', ')}${others > 0 ? ' ' + format(strings.AndOthers, others) : ''}`;
}

export const LikeButton: React.FC<ILikeButtonProps> = props => {
  const { service } = props;
  const [state, setState] = React.useState<ILikeState | undefined>(undefined);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [pop, setPop] = React.useState(false);

  React.useEffect(() => {
    let cancelled = false;
    setState(undefined);
    setError(undefined);
    if (!service) {
      return undefined;
    }
    service
      .getState()
      .then(s => !cancelled && setState(s))
      .catch(e => {
        console.error('[LikeWebPart] load failed', e);
        if (!cancelled) {
          setState({ count: 0, likedByMe: false, likers: [] });
          setError(strings.LoadError);
        }
      });
    return () => {
      cancelled = true;
    };
  }, [service]);

  const onClick = React.useCallback(async (): Promise<void> => {
    if (!service || !state || busy) {
      return;
    }
    const liking = !state.likedByMe;
    if (!liking && !props.allowUnlike) {
      return;
    }

    // Optimistic update so the click feels instant.
    const previous = state;
    const optimistic: ILikeState = {
      count: Math.max(0, state.count + (liking ? 1 : -1)),
      likedByMe: liking,
      likers: liking
        ? [props.currentUserName, ...state.likers.filter(n => n !== props.currentUserName)]
        : state.likers.filter(n => n !== props.currentUserName)
    };
    setState(optimistic);
    setPop(liking);
    setBusy(true);
    setError(undefined);

    try {
      const fresh = liking ? await service.like() : await service.unlike();
      // If the server hasn't caught up yet, keep the optimistic view.
      setState(fresh.likedByMe === liking ? fresh : optimistic);
    } catch (e) {
      console.error('[LikeWebPart] save failed', e);
      setState(previous);
      setError(
        e instanceof LikeServiceError && (e.status === 401 || e.status === 403)
          ? format(strings.AccessDenied, props.listTitle)
          : strings.SaveError
      );
    } finally {
      setBusy(false);
    }
  }, [service, state, busy, props.allowUnlike, props.currentUserName, props.listTitle]);

  const liked = !!state?.likedByMe;
  const text = liked && props.likedText ? props.likedText : props.displayText;
  const count = state?.count ?? 0;
  const countLabel = count === 1 ? props.countLabelSingular : props.countLabelPlural;
  const lockedLike = liked && !props.allowUnlike;

  const rootStyle = {
    ...props.themeVars,
    ...(props.accentColor ? { '--like-accent': props.accentColor } : {})
  } as React.CSSProperties;

  const ariaLabel = [
    liked ? strings.UnlikeAriaLabel : strings.LikeAriaLabel,
    props.showCount && state ? `${count} ${countLabel}`.trim() : ''
  ]
    .filter(Boolean)
    .join(', ');

  return (
    <div
      className={cx(
        styles.likeButton,
        cls[props.size],
        cls[props.alignment],
        props.layout === 'stacked' && styles.stacked
      )}
      style={rootStyle}
    >
      {text && <span className={styles.text}>{text}</span>}

      {!service ? (
        <span className={styles.message}>{props.unavailableMessage}</span>
      ) : (
        <span className={styles.controls}>
          <button
            type="button"
            className={cx(styles.button, cls[props.buttonStyle], liked && styles.liked, busy && styles.busy)}
            onClick={onClick}
            disabled={!state || lockedLike}
            aria-pressed={liked}
            aria-busy={busy || !state}
            aria-label={ariaLabel}
            title={props.showLikers && state ? likersTooltip(state.likers, count) : undefined}
          >
            <LikeIcon
              iconType={props.iconType}
              emoji={props.customEmoji}
              liked={liked}
              className={cx(
                styles.icon,
                props.iconType === 'emoji' && !liked && styles.emojiOff,
                pop && styles.pop
              )}
              onAnimationEnd={() => setPop(false)}
            />
            {props.showCount && (
              <span className={styles.count} aria-live="polite">
                {state ? `${formatCount(count)}${countLabel ? ' ' + countLabel : ''}` : strings.Loading}
              </span>
            )}
          </button>
          {error && (
            <span className={styles.error} role="alert">
              {error}
            </span>
          )}
        </span>
      )}
    </div>
  );
};
