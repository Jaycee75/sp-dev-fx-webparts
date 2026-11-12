import * as React from 'react';
import styles from './PageRollup.module.scss';
import { AspectRatio, HoverEffect, IRollupItem, IRollupQuery, SizeMode, TitleAlign, TitleSize } from '../models/IPageRollupProps';
import { RollupHeading } from './RollupHeading';
import { RollupService } from '../services/RollupService';
import { IRollupDisplay, RollupTile } from './RollupTile';
import { ICarouselOptions, SocialCarousel, SpotlightCarousel, TickerCarousel } from './Carousels';

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
  carousel: ICarouselOptions;
  title: string;
  titleSize: TitleSize;
  titleAlign: TitleAlign;
  titlePlaceholder: string;
  onTitleChange: (title: string) => void;
  onTitleFormatChange: (size: TitleSize, align: TitleAlign) => void;
  dateLabel: string;
  isConfigured: boolean;
  /** True while the web part is still working out which library to use (shows the loading skeleton). */
  pending?: boolean;
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
    // Repeating mix of wide and tall tiles (no single feature tile, unlike Hero)
    if (cols >= 3) {
      const p = index % 6;
      if (p === 0 || p === 4) return { c: 2, r: 1 };
      if (p === 1 || p === 3) return { c: 1, r: 2 };
    } else if (cols === 2 && index % 3 === 0) {
      return { c: 2, r: 1 };
    }
  }
  return { c: 1, r: 1 };
}

/** Add ?rollupdebug=1 to the page address to see load timings in the browser console (F12). */
function logTimings(timings: Record<string, number>, query: IRollupQuery, count: number): void {
  try {
    if (window.location.search.indexOf('rollupdebug=1') === -1) return;
    console.info(`[PageRollup] ${count} items (${query.listId}, layout data) timings ms:`, timings);
  } catch {
    // ignore
  }
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
    // Batch to one update per frame, and ignore small changes such as a page scrollbar appearing,
    // which could otherwise make the layout flip back and forth.
    let frame = 0;
    const ro = new ResizeObserver((entries) => {
      const w = Math.floor(entries[0].contentRect.width);
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => setWidth((prev) => (prev === 0 || Math.abs(prev - w) > 20 ? w : prev)));
    });
    ro.observe(el);
    return () => {
      cancelAnimationFrame(frame);
      ro.disconnect();
    };
  }, []);
  return width;
}

export const PageRollup: React.FC<IPageRollupProps> = (props) => {
  const { service, query, display, size, hoverEffect, carousel, title, dateLabel, isConfigured, isEditMode, onConfigure, themeVars } = props;
  const rootRef = React.useRef<HTMLDivElement>(null);
  const stripRef = React.useRef<HTMLDivElement>(null);
  const width = useWidth(rootRef);

  const [items, setItems] = React.useState<IRollupItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');
  const [stripPos, setStripPos] = React.useState<{ start: boolean; end: boolean }>({ start: true, end: false });

  // ----- Data ---------------------------------------------------------------
  const queryKey = JSON.stringify(query);
  const firstLoad = React.useRef(true);
  React.useEffect(() => {
    if (!isConfigured || props.pending) {
      setLoading(!!props.pending);
      return undefined;
    }
    let cancelled = false;
    setError('');

    // 1. Show the last result for these settings straight away (shared across web parts and edit/view switches)
    const hit = service.peekResult(query);
    if (hit) {
      setItems(hit.result.items);
      setLoading(false);
    } else {
      setLoading(true);
    }
    // Recent enough: no need to ask SharePoint again
    if (hit && hit.age < 60 * 1000 && !hit.result.viewsPending) {
      firstLoad.current = false;
      return undefined;
    }

    const run = (): void => {
      const t0 = performance.now();
      const timings: Record<string, number> = {};
      const mark = (name: string): void => {
        timings[name] = Math.round(performance.now() - t0);
      };
      // Previous "new" flags, so badges don't blink while they're re-checked
      const prevNew: Record<number, boolean | undefined> = {};
      (hit ? hit.result.items : []).forEach((i) => {
        prevNew[i.id] = i.isNew;
      });

      service
        .getItems(query, true)
        .then((result) => {
          if (cancelled) return;
          mark('items');
          // 2. Tiles first, with the last known view counts and "new" flags so nothing blinks...
          const first = service.knownViews(query, result.items).map((i) => ({ ...i, isNew: prevNew[i.id] }));
          setItems(first);
          setLoading(false);
          service.storeResult(query, { ...result, items: first });
          // 3. ...then fresh view counts and "new" flags. Each fills in as soon as it arrives, so a slow
          //    view-count source never holds up the NEW badges (or the other way round).
          let latest: IRollupItem[] = first;
          let viewsComplete = !result.viewsPending;
          const apply = (update: (cur: IRollupItem[]) => IRollupItem[]): void => {
            if (cancelled) return;
            latest = update(latest);
            setItems(latest);
            service.storeResult(query, { ...result, items: latest, viewsPending: result.viewsPending && !viewsComplete });
          };
          const jobs: Promise<void>[] = [];
          if (result.viewsPending && first.length > 0) {
            jobs.push(
              service.loadViews(query, first).then((views) => {
                mark(`views (${views.source})`);
                viewsComplete = views.complete;
                const byId: Record<number, IRollupItem> = {};
                views.items.forEach((i) => {
                  byId[i.id] = i;
                });
                apply((cur) =>
                  cur.map((i) => (byId[i.id] ? { ...i, viewsLifetime: byId[i.id].viewsLifetime, viewsRecent: byId[i.id].viewsRecent } : i))
                );
              })
            );
          }
          if (query.newDays > 0 && first.length > 0) {
            jobs.push(
              service
                .loadNewFlags(query, first)
                .then((withNew) => {
                  mark('new');
                  const flags: Record<number, boolean | undefined> = {};
                  withNew.forEach((i) => {
                    flags[i.id] = i.isNew;
                  });
                  apply((cur) => cur.map((i) => ({ ...i, isNew: flags[i.id] })));
                })
                .catch(() => undefined)
            );
          }
          Promise.all(jobs)
            .then(() => logTimings(timings, query, latest.length))
            .catch(() => undefined);
        })
        .catch((e: Error) => {
          if (!cancelled) {
            setError(e.message || String(e));
            if (!hit) setItems([]);
            setLoading(false);
          }
        });
    };

    // First load runs immediately; later changes (typing in the property pane) wait a moment
    let timer = 0;
    if (firstLoad.current) {
      firstLoad.current = false;
      run();
    } else {
      timer = window.setTimeout(run, 400);
    }
    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [queryKey, isConfigured, props.pending]);

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
  const header = (
    <RollupHeading
      title={title}
      size={props.titleSize}
      align={props.titleAlign}
      chip={dateLabel}
      isEditMode={isEditMode}
      placeholder={props.titlePlaceholder}
      onTitleChange={props.onTitleChange}
      onFormatChange={props.onTitleFormatChange}
    />
  );

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

  if (!isConfigured && !props.pending) {
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
  } else if (layout === 'spotlight' || layout === 'social' || layout === 'ticker') {
    if (loading) {
      content = <div className={styles.skeleton} style={{ height: fixed ? fixedH : Math.max(200, Math.min(520, w * ratio * (layout === 'ticker' ? 0.9 : 0.6))) }} />;
    } else {
      const common = { items, display, options: carousel, width: w, gap, ratio, fixedHeight: fixed ? fixedH : undefined };
      content =
        layout === 'spotlight' ? (
          <SpotlightCarousel {...common} />
        ) : layout === 'social' ? (
          <SocialCarousel {...common} visible={cols} />
        ) : (
          <TickerCarousel {...common} />
        );
    }
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
                  <RollupTile item={item} display={display} height={tileH} imgWidth={tileW} />
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
                  <RollupTile item={item} display={display} height={rowH} imgWidth={(w - gap * (cols - 1)) / cols} />
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
                      imgWidth={tileW * s.c + gap * (s.c - 1)}
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
