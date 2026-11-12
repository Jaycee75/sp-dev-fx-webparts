import * as React from 'react';
import styles from './PageRollup.module.scss';
import { ImageQuality, IRollupItem, OverlayStyle, RollupLayout, ViewsPeriod } from '../models/IPageRollupProps';
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
  imageQuality: ImageQuality;
}

export interface IRollupTileProps {
  item: IRollupItem;
  display: IRollupDisplay;
  /** Fixed tile height in px (fixed-size mode, list rows and filmstrip). */
  height?: number;
  /** Media height as a % of width (padding-top trick) for auto-height cards. */
  mediaRatio?: number;
  large?: boolean;
  /** Width the picture is drawn at, in CSS pixels; used to request a sharp enough image. */
  imgWidth?: number;
  style?: React.CSSProperties;
}

export const EyeIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8S1 12 1 12z" />
    <circle cx="12" cy="12" r="3" />
  </svg>
);

export function hashHue(text: string): number {
  let h = 0;
  for (let i = 0; i < text.length; i++) {
    h = (h * 31 + text.charCodeAt(i)) % 360;
  }
  return h;
}

export const Placeholder: React.FC<{ item: IRollupItem }> = ({ item }) => {
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

/** SharePoint preview handler sizes (approximate longest edge in px) for its resolution parameter. */
const PREVIEW_SIZES = [300, 480, 750, 1024, 1600, 2560];

/**
 * Picture sources in the order to try them. "sharp" prefers the full-size image (the browser scales it down);
 * "fast" prefers a SharePoint-resized copy that is just big enough for the tile.
 */
export function pickImageSources(item: IRollupItem, quality: ImageQuality, cssWidth: number): string[] {
  const dpr = typeof window !== 'undefined' && window.devicePixelRatio ? Math.min(window.devicePixelRatio, 2) : 1;
  const px = Math.max(200, Math.round((cssWidth || 400) * dpr));
  let res = PREVIEW_SIZES.findIndex((size) => size >= px);
  if (res === -1) res = PREVIEW_SIZES.length - 1;
  const preview = item.previewBase ? `${item.previewBase}&resolution=${res}` : '';
  const thumb = item.thumbTemplate ? item.thumbTemplate.replace('{W}', String(Math.ceil(px / 100) * 100)) : '';

  const order = quality === 'fast' ? [thumb, preview, item.originalImage] : [item.originalImage, thumb, preview];
  return order.filter((s, i, a) => !!s && a.indexOf(s) === i);
}

export const TileImage: React.FC<{ item: IRollupItem; quality: ImageQuality; width?: number }> = ({ item, quality, width }) => {
  // Bucket the width so resizing the page doesn't keep swapping images
  const bucket = Math.ceil((width || 400) / 200) * 200;
  const sources = React.useMemo(() => pickImageSources(item, quality, bucket), [item, quality, bucket]);
  const [index, setIndex] = React.useState(0);
  React.useEffect(() => setIndex(0), [sources.join('|')]);

  if (index >= sources.length) {
    return <Placeholder item={item} />;
  }
  return <img src={sources[index]} alt="" loading="lazy" decoding="async" draggable={false} onError={() => setIndex(index + 1)} />;
};

export const RollupTile: React.FC<IRollupTileProps> = ({ item, display, height, mediaRatio, large, imgWidth, style }) => {
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
        <TileImage item={item} quality={display.imageQuality} width={isList ? Math.min(220, (imgWidth || 300) * 0.3) : imgWidth} />
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
