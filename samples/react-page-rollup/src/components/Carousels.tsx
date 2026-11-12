import * as React from 'react';
import styles from './Carousels.module.scss';
import { CarouselBackground, ImageQuality, IRollupItem } from '../models/IPageRollupProps';
import { EyeIcon, hashHue, IRollupDisplay, pickImageSources, RollupTile, TileImage } from './RollupTile';
import { formatCount } from '../utils/fieldValues';

export interface ICarouselOptions {
  autoplay: boolean;
  /** Seconds each slide stays before moving on. */
  interval: number;
  /** Slide animation length in ms. */
  speed: number;
  pauseOnHover: boolean;
  loop: boolean;
  arrows: boolean;
  dots: boolean;
  /** Spotlight: width of the middle slide as a % of the web part. */
  slideWidth: number;
  background: CarouselBackground;
  bgColor: string;
  bgColor2: string;
  /** Ticker: how many rows are visible at once. */
  tickerRows: number;
}

interface ICarouselBaseProps {
  items: IRollupItem[];
  display: IRollupDisplay;
  options: ICarouselOptions;
  width: number;
  gap: number;
  ratio: number;
  fixedHeight?: number;
}

type Axis = 'x' | 'y';

// ---------------------------------------------------------------------------
// Shared carousel engine: index, autoplay, pause, drag/swipe, keyboard
// ---------------------------------------------------------------------------

function mod(n: number, m: number): number {
  return ((n % m) + m) % m;
}

interface IEngine {
  index: number;
  go: (n: number) => void;
  next: () => void;
  prev: () => void;
  drag: number;
  isDragging: boolean;
  /** Relative position of item i to the current one, wrapped when looping. */
  rel: (i: number, lowest: number) => number;
  /** True when an item jumped across the loop seam and must not animate. */
  jumped: (i: number, r: number) => boolean;
  rootProps: React.HTMLAttributes<HTMLDivElement>;
  stageProps: React.HTMLAttributes<HTMLDivElement>;
}

function useCarouselEngine(count: number, maxIndex: number, options: ICarouselOptions, axis: Axis, stepPx: number): IEngine {
  const loop = options.loop && maxIndex > 0;
  const [index, setIndex] = React.useState(0);
  const [hovered, setHovered] = React.useState(false);
  const [focused, setFocused] = React.useState(false);
  const [drag, setDrag] = React.useState(0);
  const [isDragging, setDragging] = React.useState(false);
  const start = React.useRef<{ x: number; y: number; id: number } | null>(null);
  const moved = React.useRef(false);
  const lastRel = React.useRef<Record<number, number>>({});

  const go = React.useCallback(
    (n: number) => setIndex(loop ? mod(n, count) : Math.max(0, Math.min(maxIndex, n))),
    [loop, count, maxIndex]
  );
  const next = React.useCallback(() => setIndex((i) => (loop ? mod(i + 1, count) : i >= maxIndex ? 0 : i + 1)), [loop, count, maxIndex]);
  const prev = React.useCallback(() => setIndex((i) => (loop ? mod(i - 1, count) : i <= 0 ? maxIndex : i - 1)), [loop, count, maxIndex]);

  // Keep index valid when the item count changes
  React.useEffect(() => {
    setIndex((i) => (i > Math.max(0, loop ? count - 1 : maxIndex) ? 0 : i));
  }, [count, maxIndex, loop]);

  // Autoplay
  const paused = isDragging || (options.pauseOnHover && (hovered || focused));
  React.useEffect(() => {
    if (!options.autoplay || paused || maxIndex <= 0) return undefined;
    const ms = Math.max(1000, (options.interval || 5) * 1000);
    const timer = window.setInterval(() => {
      if (typeof document !== 'undefined' && document.hidden) return;
      next();
    }, ms);
    return () => window.clearInterval(timer);
  }, [options.autoplay, options.interval, paused, maxIndex, next]);

  // Remember positions after each render so wrap-around jumps can skip their animation
  const rels: Record<number, number> = {};
  const rel = (i: number, lowest: number): number => {
    const r = loop ? mod(i - index - lowest, count) + lowest : i - index;
    rels[i] = r;
    return r;
  };
  const jumped = (i: number, r: number): boolean => {
    const before = lastRel.current[i];
    return before !== undefined && Math.abs(before - r) > 1.5;
  };
  React.useEffect(() => {
    lastRel.current = { ...lastRel.current, ...rels };
  });

  const onPointerDown = (e: React.PointerEvent<HTMLDivElement>): void => {
    if (e.pointerType === 'mouse' && e.button !== 0) return;
    start.current = { x: e.clientX, y: e.clientY, id: e.pointerId };
    moved.current = false;
  };
  const onPointerMove = (e: React.PointerEvent<HTMLDivElement>): void => {
    if (!start.current || start.current.id !== e.pointerId) return;
    const d = axis === 'x' ? e.clientX - start.current.x : e.clientY - start.current.y;
    const other = axis === 'x' ? e.clientY - start.current.y : e.clientX - start.current.x;
    if (!moved.current) {
      if (Math.abs(d) < 6 || Math.abs(other) > Math.abs(d)) return; // let the page scroll on the other axis
      moved.current = true;
      setDragging(true);
      try {
        e.currentTarget.setPointerCapture(e.pointerId);
      } catch {
        // ignore
      }
    }
    // Resist at the ends when not looping
    const atEdge = !loop && ((index === 0 && d > 0) || (index >= maxIndex && d < 0));
    setDrag(atEdge ? d / 3 : d);
  };
  const endDrag = (): void => {
    if (!start.current) return;
    const d = drag;
    start.current = null;
    setDragging(false);
    setDrag(0);
    if (moved.current) {
      const threshold = Math.min(80, Math.max(30, stepPx * 0.15));
      if (d <= -threshold) next();
      else if (d >= threshold) prev();
    }
  };
  const onClickCapture = (e: React.MouseEvent): void => {
    if (moved.current) {
      e.preventDefault();
      e.stopPropagation();
      moved.current = false;
    }
  };
  const onKeyDown = (e: React.KeyboardEvent): void => {
    const back = axis === 'x' ? 'ArrowLeft' : 'ArrowUp';
    const fwd = axis === 'x' ? 'ArrowRight' : 'ArrowDown';
    if (e.key === back) {
      e.preventDefault();
      prev();
    } else if (e.key === fwd) {
      e.preventDefault();
      next();
    }
  };

  return {
    index,
    go,
    next,
    prev,
    drag,
    isDragging,
    rel,
    jumped,
    rootProps: {
      onMouseEnter: () => setHovered(true),
      onMouseLeave: () => setHovered(false),
      onFocus: () => setFocused(true),
      onBlur: (e: React.FocusEvent<HTMLDivElement>) => {
        if (!e.currentTarget.contains(e.relatedTarget as Node)) setFocused(false);
      },
      onKeyDown
    },
    stageProps: {
      onPointerDown,
      onPointerMove,
      onPointerUp: endDrag,
      onPointerCancel: endDrag,
      onClickCapture,
      // Stop the browser's own link/image dragging from hijacking a swipe
      onDragStart: (e: React.DragEvent) => e.preventDefault(),
      style: { touchAction: axis === 'x' ? 'pan-y' : 'pan-x' }
    }
  };
}

const Arrow: React.FC<{ dir: 'left' | 'right' | 'up' | 'down' }> = ({ dir }) => {
  const d = { left: 'M15 18l-6-6 6-6', right: 'M9 18l6-6-6-6', up: 'M18 15l-6-6-6 6', down: 'M6 9l6 6 6-6' }[dir];
  return (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d={d} />
    </svg>
  );
};

const Dots: React.FC<{ count: number; index: number; go: (i: number) => void; light: boolean }> = ({ count, index, go, light }) => {
  if (count <= 1 || count > 20) return null;
  return (
    <div className={`${styles.dots} ${light ? styles.dotsLight : ''}`} role="tablist">
      {Array.from({ length: count }).map((_, i) => (
        <button
          key={i}
          type="button"
          role="tab"
          aria-selected={i === index}
          aria-label={`Go to item ${i + 1}`}
          className={`${styles.dot} ${i === index ? styles.dotActive : ''}`}
          onClick={() => go(i)}
        />
      ))}
    </div>
  );
};

function transition(ms: number, skip: boolean, dragging: boolean): string {
  if (skip || dragging) return 'none';
  const t = `${Math.max(100, ms)}ms cubic-bezier(0.22, 0.8, 0.2, 1)`;
  return `transform ${t}, opacity ${t}, filter ${t}`;
}

// ---------------------------------------------------------------------------
// 1. Spotlight: the current slide in the middle, neighbours peeking at the sides,
//    on a coloured / gradient / blurred-picture background.
// ---------------------------------------------------------------------------

export const SpotlightCarousel: React.FC<ICarouselBaseProps> = ({ items, display, options, width, gap, ratio, fixedHeight }) => {
  const n = items.length;
  // On narrow screens (phones, side columns) the middle slide takes most of the width
  const pct = Math.max(30, Math.min(95, options.slideWidth || 60));
  const slideW = Math.max(160, Math.round((width * (width < 600 ? Math.max(pct, 82) : pct)) / 100));
  const sideScale = 0.8;
  const padY = options.background === 'none' ? 8 : 32;
  const dotsH = options.dots && n > 1 ? 28 : 0;
  const slideH = fixedHeight
    ? Math.max(120, fixedHeight - padY * 2 - dotsH)
    : Math.max(140, Math.min(640, Math.round(slideW * ratio)));
  const step = slideW / 2 + gap + (slideW * sideScale) / 2;

  const engine = useCarouselEngine(n, n - 1, options, 'x', step);
  const current = items[engine.index];
  const lowest = -Math.floor((n - 1) / 2);
  const light = options.background !== 'none' && options.background !== 'theme';

  let bgStyle: React.CSSProperties = {};
  if (options.background === 'theme') bgStyle = { background: 'var(--pr-primary-soft)' };
  if (options.background === 'color') bgStyle = { background: options.bgColor || '#1b1f3b' };
  if (options.background === 'gradient')
    bgStyle = { background: `linear-gradient(135deg, ${options.bgColor || '#1b1f3b'} 0%, ${options.bgColor2 || '#3a1c71'} 100%)` };
  if (options.background === 'blur') bgStyle = { background: '#1a1a1a' };

  const blurSrc = options.background === 'blur' && current ? pickImageSources(current, 'fast', 400)[0] : '';

  return (
    <div
      className={`${styles.carousel} ${styles.spotlight}`}
      style={{ ...bgStyle, paddingTop: padY, paddingBottom: padY }}
      aria-roledescription="carousel"
      tabIndex={0}
      {...engine.rootProps}
    >
      {options.background === 'blur' && (
        <div className={styles.blurBg} aria-hidden="true">
          {blurSrc && <img key={blurSrc} src={blurSrc} alt="" />}
        </div>
      )}
      <div className={styles.stage} {...engine.stageProps} style={{ height: slideH, ...engine.stageProps.style }}>
        {items.map((item, i) => {
          const r = engine.rel(i, lowest);
          const abs = Math.abs(r);
          const x = r === 0 ? 0 : Math.sign(r) * (step + (abs - 1) * (slideW * sideScale + gap));
          const hidden = abs > 2;
          const isCurrent = r === 0;
          return (
            <div
              key={item.id}
              className={styles.spotSlide}
              aria-hidden={!isCurrent}
              style={{
                width: slideW,
                height: slideH,
                transform: `translateX(calc(-50% + ${x + engine.drag}px)) scale(${isCurrent ? 1 : sideScale})`,
                opacity: hidden ? 0 : 1,
                filter: isCurrent ? 'none' : 'brightness(0.55) saturate(0.8)',
                zIndex: 10 - abs,
                pointerEvents: hidden ? 'none' : undefined,
                transition: transition(options.speed, engine.jumped(i, r), engine.isDragging)
              }}
              onClickCapture={(e) => {
                if (!isCurrent) {
                  // Clicking a side slide brings it to the middle instead of opening it
                  e.preventDefault();
                  e.stopPropagation();
                  engine.go(i);
                }
              }}
            >
              <RollupTile item={item} display={{ ...display, layout: 'tiles' }} height={slideH} large={isCurrent && slideW > 420} imgWidth={slideW} />
            </div>
          );
        })}
      </div>
      {options.arrows && n > 1 && (
        <>
          <button type="button" className={`${styles.arrow} ${styles.arrowLeft}`} onClick={engine.prev} aria-label="Previous">
            <Arrow dir="left" />
          </button>
          <button type="button" className={`${styles.arrow} ${styles.arrowRight}`} onClick={engine.next} aria-label="Next">
            <Arrow dir="right" />
          </button>
        </>
      )}
      {options.dots && <Dots count={n} index={engine.index} go={engine.go} light={light} />}
    </div>
  );
};

// ---------------------------------------------------------------------------
// 2. Social cards: Instagram-style post cards sliding along one at a time.
// ---------------------------------------------------------------------------

const SocialCard: React.FC<{ item: IRollupItem; display: IRollupDisplay; width: number; mediaH: number; height: number; quality: ImageQuality }> = ({
  item,
  display,
  width,
  mediaH,
  height,
  quality
}) => {
  const [headingField, ...metaFields] = display.overlayFields.length > 0 ? display.overlayFields : ['Title'];
  const heading = item.values[headingField] || item.title;
  const meta = metaFields.map((f) => item.values[f]).filter((v) => !!v);
  // Column 2 is the card header (e.g. author), column 3 the line under it (e.g. date), the rest go in the caption
  const who = meta[0] || '';
  const when = meta[1] || '';
  const rest = meta.slice(2);
  const views = display.viewsPeriod === 'recent' ? item.viewsRecent : item.viewsLifetime;
  const hue = hashHue(who || heading);
  const initials = (who || heading)
    .split(/\s+/)
    .filter((w) => /[A-Za-z0-9]/.test(w))
    .slice(0, 2)
    .map((w) => w.charAt(0).toUpperCase())
    .join('');

  return (
    <a
      className={styles.social}
      href={item.url}
      target={display.openInNewTab ? '_blank' : undefined}
      rel={display.openInNewTab ? 'noopener noreferrer' : undefined}
      data-interception={display.openInNewTab ? 'off' : undefined}
      aria-label={item.title}
      style={{ width, height }}
      draggable={false}
    >
      {who && (
        <div className={styles.socialHead}>
          <span className={styles.avatar} style={{ background: `linear-gradient(135deg, hsl(${hue},60%,50%), hsl(${(hue + 50) % 360},65%,40%))` }}>
            {initials}
          </span>
          <span className={styles.socialWho}>
            <b>{who}</b>
            {when && <small>{when}</small>}
          </span>
        </div>
      )}
      <div className={styles.socialMedia} style={{ height: mediaH }}>
        <TileImage item={item} quality={quality} width={width} />
      </div>
      <div className={styles.socialFoot}>
        {display.showViews && typeof views === 'number' && (
          <div className={styles.socialStats}>
            <EyeIcon />
            <span>
              <b>{formatCount(views)}</b> views
            </span>
          </div>
        )}
        {display.showOverlay && (
          <p className={styles.socialCaption}>
            <b>{heading}</b>
            {rest.length > 0 && <span> {rest.join(' · ')}</span>}
          </p>
        )}
      </div>
    </a>
  );
};

export const SocialCarousel: React.FC<ICarouselBaseProps & { visible: number }> = ({ items, display, options, width, gap, ratio, fixedHeight, visible }) => {
  const n = items.length;
  const shown = Math.max(1, Math.min(visible, n));
  const cardW = (width - gap * (shown - 1)) / shown;
  const headH = 56;
  const footH = (display.showViews ? 30 : 0) + (display.showOverlay ? 52 : 0) + 16;
  const dotsH = options.dots && n > shown ? 28 : 0;
  const cardH = fixedHeight ? Math.max(180, fixedHeight - dotsH - 8) : Math.round(headH + cardW * ratio + footH);
  const mediaH = Math.max(60, cardH - headH - footH);
  const maxIndex = Math.max(0, n - shown);

  const engine = useCarouselEngine(n, maxIndex, options, 'x', cardW + gap);
  const loop = options.loop && maxIndex > 0;

  return (
    <div className={`${styles.carousel} ${styles.socialWrap}`} aria-roledescription="carousel" tabIndex={0} {...engine.rootProps}>
      <div className={styles.stage} {...engine.stageProps} style={{ height: cardH + 8, ...engine.stageProps.style }}>
        {items.map((item, i) => {
          const r = engine.rel(i, loop ? -1 : -n);
          const hidden = r < -1 || r > shown;
          return (
            <div
              key={item.id}
              className={styles.track}
              aria-hidden={r < 0 || r >= shown}
              style={{
                width: cardW,
                transform: `translateX(${r * (cardW + gap) + engine.drag}px)`,
                opacity: r < 0 || r >= shown ? 0 : 1,
                visibility: hidden ? 'hidden' : 'visible',
                transition: transition(options.speed, engine.jumped(i, r), engine.isDragging)
              }}
            >
              <SocialCard item={item} display={display} width={cardW} mediaH={mediaH} height={cardH} quality={display.imageQuality} />
            </div>
          );
        })}
      </div>
      {options.arrows && maxIndex > 0 && (
        <>
          <button type="button" className={`${styles.arrow} ${styles.arrowLeft} ${styles.arrowSmall}`} onClick={engine.prev} aria-label="Previous">
            <Arrow dir="left" />
          </button>
          <button type="button" className={`${styles.arrow} ${styles.arrowRight} ${styles.arrowSmall}`} onClick={engine.next} aria-label="Next">
            <Arrow dir="right" />
          </button>
        </>
      )}
      {options.dots && maxIndex > 0 && <Dots count={loop ? n : maxIndex + 1} index={engine.index} go={engine.go} light={false} />}
    </div>
  );
};

// ---------------------------------------------------------------------------
// 3. Slide-up ticker: a column of tiles; the top one slides up and away and the
//    next one rises in from the bottom.
// ---------------------------------------------------------------------------

export const TickerCarousel: React.FC<ICarouselBaseProps> = ({ items, display, options, width, gap, ratio, fixedHeight }) => {
  const n = items.length;
  const rows = Math.max(1, Math.min(options.tickerRows || 3, n));
  const rowH = fixedHeight
    ? Math.max(80, (fixedHeight - gap * (rows - 1)) / rows)
    : Math.round(Math.min(220, Math.max(110, width * ratio * 0.5)));
  const maxIndex = Math.max(0, n - rows);
  const engine = useCarouselEngine(n, maxIndex, options, 'y', rowH + gap);
  const loop = options.loop && maxIndex > 0;
  const stageH = rows * rowH + (rows - 1) * gap;

  return (
    <div className={`${styles.carousel} ${styles.ticker}`} aria-roledescription="carousel" tabIndex={0} {...engine.rootProps}>
      <div className={styles.stage} {...engine.stageProps} style={{ height: stageH, ...engine.stageProps.style }}>
        {items.map((item, i) => {
          const r = engine.rel(i, loop ? -1 : -n);
          const out = r < 0 || r >= rows;
          return (
            <div
              key={item.id}
              className={styles.track}
              aria-hidden={out}
              style={{
                width: '100%',
                height: rowH,
                transform: `translateY(${r * (rowH + gap) + engine.drag}px) scale(${out ? 0.94 : 1})`,
                opacity: out ? 0 : 1,
                visibility: r < -1 || r > rows ? 'hidden' : 'visible',
                transition: transition(options.speed, engine.jumped(i, r), engine.isDragging)
              }}
            >
              <RollupTile item={item} display={{ ...display, layout: 'tiles' }} height={rowH} imgWidth={width} />
            </div>
          );
        })}
      </div>
      {options.arrows && maxIndex > 0 && (
        <div className={styles.tickerNav}>
          <button type="button" className={`${styles.arrow} ${styles.arrowSmall}`} onClick={engine.prev} aria-label="Previous">
            <Arrow dir="up" />
          </button>
          <button type="button" className={`${styles.arrow} ${styles.arrowSmall}`} onClick={engine.next} aria-label="Next">
            <Arrow dir="down" />
          </button>
        </div>
      )}
    </div>
  );
};
