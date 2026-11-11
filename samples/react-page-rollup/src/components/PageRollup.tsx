import * as React from 'react';
import styles from './PageRollup.module.scss';
import { AspectRatio, HoverEffect, IRollupItem, IRollupQuery, SizeMode } from '../models/IPageRollupProps';
import { RollupService } from '../services/RollupService';
import { IRollupDisplay, RollupTile } from './RollupTile';

export interface IRollupSize {
  sizeMode: SizeMode;
  fixedHeight: number;
  columns: number;
  tileSize: number;
  aspectRatio: AspectRatio;
  gap: number;
  cornerRadius: number;
}

export interface IPageRollupProps {
  service: RollupService;
  query: IRollupQuery;
  display: IRollupDisplay;
  size: IRollupSize;
  hoverEffect: HoverEffect;
  title: string;
  dateLabel: string;
  isConfigured: boolean;
  isEditMode: boolean;
  onConfigure: () => void;
  themeVars: Record<string, string>;
}

const RATIOS: Record<AspectRatio, number> = { '16:9': 9 / 16, '4:3': 3 / 4, '1:1': 1, '3:4': 4 / 3 };
const MIN_ROW = 100;

const Chevron: React.FC<{ dir: 'left' | 'right' }> = ({ dir }) => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d={dir === 'left' ? 'M15 18l-6-6 6-6' : 'M9 18l6-6-6-6'} />
  </svg>
);

const StackIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <rect x="3" y="3" width="8" height="8" rx="2" />
    <rect x="13" y="3" width="8" height="8" rx="2" />
    <rect x="3" y="13" width="8" height="8" rx="2" />
    <rect x="13" y="13" width="8" height="8" rx="2" />
  </svg>
);

/** Grid cell spans for the hero and mosaic layouts. */
function getSpan(layout: string, index: number, cols: number): { c: number; r: number } {
  if (layout === 'hero') {
    if (index === 0 && cols >= 3) return { c: 2, r: 2 };
    if (index === 0 && cols === 2) return { c: 2, r: 1 };
  }
  if (layout === 'mosaic') {
    if (cols >= 3) {
      const p = index % 7;
      if (p === 0) return { c: 2, r: 2 };
      if (p === 3) return { c: 1, r: 2 };
      if (p === 5) return { c: 2, r: 1 };
    } else if (cols === 2 && index % 3 === 0) {
      return { c: 2, r: 1 };
    }
  }
  return { c: 1, r: 1 };
}

function useWidth(ref: React.RefObject<HTMLDivElement>): number {
  const [width, setWidth] = React.useState(0);
  React.useEffect(() => {
    const el = ref.current;
    if (!el) return undefined;
    setWidth(el.clientWidth);
    if (typeof ResizeObserver === 'undefined') {
      const onResize = (): void => setWidth(el.clientWidth);
      window.addEventListener('resize', onResize);
      return () => window.removeEventListener('resize', onResize);
    }
    const ro = new ResizeObserver((entries) => {
      const w = Math.floor(entries[0].contentRect.width);
      setWidth((prev) => (Math.abs(prev - w) > 1 ? w : prev));
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, []);
  return width;
}

export const PageRollup: React.FC<IPageRollupProps> = (props) => {
  const { service, query, display, size, hoverEffect, title, dateLabel, isConfigured, isEditMode, onConfigure, themeVars } = props;
  const rootRef = React.useRef<HTMLDivElement>(null);
  const stripRef = React.useRef<HTMLDivElement>(null);
  const width = useWidth(rootRef);

  const [items, setItems] = React.useState<IRollupItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');
  const [stripPos, setStripPos] = React.useState<{ start: boolean; end: boolean }>({ start: true, end: false });

  // ----- Data ---------------------------------------------------------------
  const queryKey = JSON.stringify(query);
  React.useEffect(() => {
    if (!isConfigured) {
      setLoading(false);
      return undefined;
    }
    let cancelled = false;
    setLoading(true);
    setError('');
    // Small delay so typing in the property pane doesn't fire a request per keystroke
    const timer = window.setTimeout(() => {
      service
        .getItems(query)
        .then((result) => {
          if (!cancelled) {
            setItems(result.items);
            setLoading(false);
          }
        })
        .catch((e: Error) => {
          if (!cancelled) {
            setError(e.message || String(e));
            setItems([]);
            setLoading(false);
          }
        });
    }, 400);
    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [queryKey, isConfigured]);

  // ----- Geometry -------------------------------------------------------------
  const layout = display.layout;
  const gap = Math.max(0, size.gap);
  const w = width || 800;
  const ratio = RATIOS[size.aspectRatio] || RATIOS['16:9'];
  const fixed = size.sizeMode === 'fixed';
  const fixedH = Math.max(120, size.fixedHeight || 400);
  const count = loading ? Math.max(1, Math.min(query.maxItems, 8)) : items.length;

  let cols: number;
  if (layout === 'list') {
    cols = size.columns > 0 ? Math.min(size.columns, 3) : w >= 1000 ? 2 : 1;
  } else {
    cols = size.columns > 0 ? size.columns : Math.max(1, Math.floor((w + gap) / (Math.max(120, size.tileSize) + gap)));
  }
  cols = Math.max(1, Math.min(cols, 8));
  const tileW = (w - gap * (cols - 1)) / cols;

  // Rows needed (hero/mosaic tiles take more than one cell)
  let cells = 0;
  for (let i = 0; i < count; i++) {
    const s = getSpan(layout, i, cols);
    cells += s.c * s.r;
  }
  const rows = Math.max(1, Math.ceil(cells / cols));

  let rowH: number;
  if (layout === 'list') {
    rowH = fixed ? Math.max(72, (fixedH - gap * 0.75 * (Math.ceil(count / cols) - 1)) / Math.ceil(count / cols)) : Math.max(96, Math.min(160, size.tileSize * 0.45));
  } else if (fixed && layout !== 'filmstrip') {
    rowH = Math.max(MIN_ROW, (fixedH - gap * (rows - 1)) / rows);
  } else {
    rowH = tileW * ratio;
  }
  const overflowsFixed = fixed && layout !== 'filmstrip' && rowH * rows + gap * (rows - 1) > fixedH + 1;

  const rootStyle = {
    ...themeVars,
    '--pr-gap': `${gap}px`,
    '--pr-radius': `${Math.max(0, size.cornerRadius)}px`
  } as React.CSSProperties;

  const hoverClass =
    hoverEffect === 'zoom' ? styles.hoverZoom : hoverEffect === 'lift' ? styles.hoverLift : hoverEffect === 'reveal' ? styles.hoverReveal : '';

  // ----- Filmstrip navigation ---------------------------------------------------
  const updateStrip = React.useCallback(() => {
    const el = stripRef.current;
    if (!el) return;
    setStripPos({ start: el.scrollLeft <= 4, end: el.scrollLeft + el.clientWidth >= el.scrollWidth - 4 });
  }, []);
  React.useEffect(updateStrip, [items, w, cols, layout]);
  const scrollStrip = (dir: number): void => {
    const el = stripRef.current;
    if (el) el.scrollBy({ left: dir * (tileW + gap) * Math.max(1, cols - (cols > 2 ? 1 : 0)), behavior: 'smooth' });
  };

  // ----- Render helpers -----------------------------------------------------------
  const header =
    title || dateLabel ? (
      <div className={styles.header}>
        {title && <h2 className={styles.heading}>{title}</h2>}
        {dateLabel && <span className={styles.chip}>{dateLabel}</span>}
      </div>
    ) : null;

  const renderState = (heading: string, text: string, isError?: boolean, button?: boolean): JSX.Element => (
    <div className={`${styles.state} ${isError ? styles.error : ''}`} role={isError ? 'alert' : 'status'}>
      <StackIcon />
      <h3>{heading}</h3>
      <p>{text}</p>
      {button && (
        <button type="button" className={styles.button} onClick={onConfigure}>
          Configure
        </button>
      )}
    </div>
  );

  let content: JSX.Element;

  if (!isConfigured) {
    content = isEditMode
      ? renderState('Page Rollup', 'Choose a site and library to roll up, then pick filters, a layout and the columns to show on each tile.', false, true)
      : <></>;
  } else if (error) {
    content = renderState("Couldn't load content", error, true, isEditMode);
  } else if (!loading && items.length === 0) {
    content = renderState(
      'Nothing to show yet',
      'No items match these settings. Try widening the date range or checking the filter values.',
      false,
      isEditMode
    );
  } else if (layout === 'filmstrip') {
    const tileH = fixed ? fixedH - 12 : Math.max(MIN_ROW, tileW * ratio);
    content = (
      <div className={styles.viewport}>
        <button type="button" className={`${styles.navButton} ${styles.navPrev}`} onClick={() => scrollStrip(-1)} disabled={stripPos.start} aria-label="Previous">
          <Chevron dir="left" />
        </button>
        <div
          ref={stripRef}
          className={styles.strip}
          style={{ gridAutoColumns: `${tileW}px` }}
          onScroll={updateStrip}
          role="list"
        >
          {loading
            ? Array.from({ length: Math.min(count, cols + 1) }).map((_, i) => <div key={i} className={styles.skeleton} style={{ height: tileH }} />)
            : items.map((item) => (
                <div key={item.id} role="listitem">
                  <RollupTile item={item} display={display} height={tileH} />
                </div>
              ))}
        </div>
        <button type="button" className={`${styles.navButton} ${styles.navNext}`} onClick={() => scrollStrip(1)} disabled={stripPos.end} aria-label="Next">
          <Chevron dir="right" />
        </button>
      </div>
    );
  } else if (layout === 'list') {
    content = (
      <div className={`${styles.viewport} ${fixed ? styles.scrollY : ''}`} style={fixed ? { maxHeight: fixedH } : undefined}>
        <div className={styles.list} style={{ gridTemplateColumns: `repeat(${cols}, minmax(0, 1fr))` }} role="list">
          {loading
            ? Array.from({ length: count }).map((_, i) => <div key={i} className={styles.skeleton} style={{ height: rowH }} />)
            : items.map((item) => (
                <div key={item.id} role="listitem">
                  <RollupTile item={item} display={display} height={rowH} />
                </div>
              ))}
        </div>
      </div>
    );
  } else {
    const isCards = layout === 'cards';
    const autoCards = isCards && !fixed;
    content = (
      <div
        className={`${styles.viewport} ${overflowsFixed ? styles.scrollY : ''}`}
        style={fixed ? { maxHeight: fixedH } : undefined}
      >
        <div
          className={styles.grid}
          style={{
            gridTemplateColumns: `repeat(${cols}, minmax(0, 1fr))`,
            gridAutoRows: autoCards ? undefined : `${rowH}px`
          }}
          role="list"
        >
          {loading
            ? Array.from({ length: count }).map((_, i) => {
                const s = getSpan(layout, i, cols);
                return (
                  <div
                    key={i}
                    className={styles.skeleton}
                    style={{
                      gridColumn: `span ${s.c}`,
                      gridRow: `span ${s.r}`,
                      height: autoCards ? tileW * ratio + 70 : undefined
                    }}
                  />
                );
              })
            : items.map((item, i) => {
                const s = getSpan(layout, i, cols);
                return (
                  <div key={item.id} role="listitem" style={{ gridColumn: `span ${s.c}`, gridRow: `span ${s.r}`, display: 'flex', minWidth: 0 }}>
                    <RollupTile
                      item={item}
                      display={display}
                      mediaRatio={ratio * 100}
                      height={isCards && fixed ? rowH : undefined}
                      large={s.c * s.r >= 4}
                      style={{ flex: '1 1 auto' }}
                    />
                  </div>
                );
              })}
        </div>
      </div>
    );
  }

  return (
    <div ref={rootRef} className={`${styles.root} ${hoverClass}`} style={rootStyle} aria-busy={loading}>
      {isConfigured && header}
      {content}
    </div>
  );
};
