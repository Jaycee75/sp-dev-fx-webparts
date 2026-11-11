import * as React from 'react';
import styles from './PageRollup.module.scss';
import { IRollupItem, OverlayStyle, RollupLayout, ViewsPeriod } from '../models/IPageRollupProps';
import { formatCount } from '../utils/fieldValues';

export interface IRollupDisplay {
  layout: RollupLayout;
  showOverlay: boolean;
  overlayStyle: OverlayStyle;
  /** Internal names of the columns shown on the tile, first one is the heading. */
  overlayFields: string[];
  fieldTitles: Record<string, string>;
  showFieldLabels: boolean;
  showViews: boolean;
  viewsPeriod: ViewsPeriod;
  openInNewTab: boolean;
}

export interface IRollupTileProps {
  item: IRollupItem;
  display: IRollupDisplay;
  /** Fixed tile height in px (fixed-size mode, list rows and filmstrip). */
  height?: number;
  /** Media height as a % of width (padding-top trick) for auto-height cards. */
  mediaRatio?: number;
  large?: boolean;
  style?: React.CSSProperties;
}

const EyeIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8S1 12 1 12z" />
    <circle cx="12" cy="12" r="3" />
  </svg>
);

function hashHue(text: string): number {
  let h = 0;
  for (let i = 0; i < text.length; i++) {
    h = (h * 31 + text.charCodeAt(i)) % 360;
  }
  return h;
}

const Placeholder: React.FC<{ item: IRollupItem }> = ({ item }) => {
  const hue = hashHue(item.title || String(item.id));
  const initials = (item.title || '?')
    .split(/\s+/)
    .filter((w) => /[A-Za-z0-9]/.test(w))
    .slice(0, 2)
    .map((w) => w.charAt(0).toUpperCase())
    .join('');
  return (
    <div
      className={styles.placeholder}
      style={{ background: `linear-gradient(135deg, hsl(${hue}, 58%, 46%) 0%, hsl(${(hue + 45) % 360}, 62%, 32%) 100%)` }}
      aria-hidden="true"
    >
      <span>{initials || '•'}</span>
      {item.fileType && item.fileType !== 'aspx' && <small>{item.fileType}</small>}
    </div>
  );
};

const TileImage: React.FC<{ item: IRollupItem }> = ({ item }) => {
  // Try the resized preview first, then the original image, then a coloured placeholder.
  const sources = [item.imageUrl, item.fallbackImageUrl].filter((s, i, a) => !!s && a.indexOf(s) === i);
  const [index, setIndex] = React.useState(0);
  React.useEffect(() => setIndex(0), [item.imageUrl, item.fallbackImageUrl]);

  if (index >= sources.length) {
    return <Placeholder item={item} />;
  }
  return <img src={sources[index]} alt="" loading="lazy" decoding="async" onError={() => setIndex(index + 1)} />;
};

export const RollupTile: React.FC<IRollupTileProps> = ({ item, display, height, mediaRatio, large, style }) => {
  const { layout, showOverlay, overlayStyle, overlayFields, fieldTitles, showFieldLabels, showViews, viewsPeriod } = display;
  const views = viewsPeriod === 'recent' ? item.viewsRecent : item.viewsLifetime;
  const isCard = layout === 'cards';
  const isList = layout === 'list';
  const textBelow = isCard || isList;

  const [headingField, ...metaFields] = overlayFields.length > 0 ? overlayFields : ['Title'];
  const heading = item.values[headingField] || (headingField === 'Title' ? item.title : '');
  const meta = metaFields
    .map((f) => ({ field: f, text: item.values[f] }))
    .filter((m) => !!m.text);

  const text = showOverlay && (heading || meta.length > 0) && (
    <div
      className={
        textBelow
          ? styles.body
          : `${styles.overlay} ${
              overlayStyle === 'glass' ? styles.overlayGlass : overlayStyle === 'solid' ? styles.overlaySolid : styles.overlayGradient
            }`
      }
    >
      {heading && <p className={`${styles.title} ${large ? styles.titleLarge : ''}`}>{heading}</p>}
      {meta.length > 0 && (
        <div className={styles.metaWrap}>
          {meta.map((m) => (
            <div key={m.field} className={styles.meta} title={m.text}>
              {showFieldLabels && <span className={styles.metaLabel}>{fieldTitles[m.field] || m.field}:</span>}
              {m.text}
            </div>
          ))}
        </div>
      )}
    </div>
  );

  // Card media keeps the chosen aspect ratio unless the tile has a fixed height.
  const mediaStyle: React.CSSProperties = isList
    ? {}
    : isCard
      ? height
        ? { flex: '1 1 auto', minHeight: '40%' }
        : { paddingTop: `${mediaRatio || 56.25}%` }
      : {};

  const tileClass = [styles.tile, isCard ? styles.cardTile : '', isList ? styles.listTile : ''].join(' ');

  return (
    <a
      className={tileClass}
      href={item.url}
      target={display.openInNewTab ? '_blank' : undefined}
      rel={display.openInNewTab ? 'noopener noreferrer' : undefined}
      data-interception={display.openInNewTab ? 'off' : undefined}
      aria-label={item.title}
      title={item.title}
      style={{ ...(height ? { height } : {}), ...(textBelow ? {} : { minHeight: height || 120 }), ...style }}
    >
      <div className={`${styles.media} ${textBelow ? '' : styles.mediaFill}`} style={mediaStyle}>
        <TileImage item={item} />
      </div>
      {showViews && typeof views === 'number' && (
        <span className={styles.views} aria-label={`${views} views`}>
          <EyeIcon />
          {formatCount(views)}
        </span>
      )}
      {text}
    </a>
  );
};
