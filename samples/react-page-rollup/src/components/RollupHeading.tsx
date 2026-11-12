import * as React from 'react';
import styles from './PageRollup.module.scss';
import { TitleAlign, TitleSize } from '../models/IPageRollupProps';

export interface IRollupHeadingProps {
  title: string;
  size: TitleSize;
  align: TitleAlign;
  /** Small pill next to the heading, e.g. "Q4 2026". */
  chip: string;
  isEditMode: boolean;
  placeholder: string;
  onTitleChange: (title: string) => void;
  onFormatChange: (size: TitleSize, align: TitleAlign) => void;
}

const SIZES: { key: TitleSize; label: string; name: string }[] = [
  { key: 'small', label: 'S', name: 'Small' },
  { key: 'medium', label: 'M', name: 'Medium' },
  { key: 'large', label: 'L', name: 'Large' },
  { key: 'xlarge', label: 'XL', name: 'Extra large' }
];

const AlignIcon: React.FC<{ align: TitleAlign }> = ({ align }) => {
  // Three lines, the shorter ones placed to show the alignment
  const x = align === 'left' ? [4, 4, 4] : align === 'center' ? [4, 7, 5.5] : [4, 10, 7];
  const w = [16, 10, 13];
  return (
    <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
      {[6, 11, 16].map((y, i) => (
        <rect key={y} x={x[i]} y={y} width={w[i]} height="2.2" rx="1.1" />
      ))}
    </svg>
  );
};

/**
 * The web part heading. On the published page it's plain text (hidden when empty).
 * In edit mode it can be typed into directly, with a small toolbar for size and alignment.
 */
export const RollupHeading: React.FC<IRollupHeadingProps> = (props) => {
  const { title, size, align, chip, isEditMode, placeholder, onTitleChange, onFormatChange } = props;
  const [text, setText] = React.useState(title || '');
  const areaRef = React.useRef<HTMLTextAreaElement>(null);

  // Pick up changes made in the property pane
  React.useEffect(() => {
    setText(title || '');
  }, [title]);

  // Grow the box to fit wrapped lines
  React.useLayoutEffect(() => {
    const el = areaRef.current;
    if (el) {
      el.style.height = 'auto';
      el.style.height = `${el.scrollHeight}px`;
    }
  }, [text, size, isEditMode]);

  const sizeClass = styles[`headingSize_${size || 'medium'}`];
  const alignStyle: React.CSSProperties = { textAlign: align || 'left' };
  const justify = align === 'center' ? 'center' : align === 'right' ? 'flex-end' : 'flex-start';

  if (!isEditMode) {
    if (!title && !chip) return null;
    return (
      <div className={styles.header} style={{ justifyContent: justify }}>
        {title && (
          <h2 className={`${styles.heading} ${sizeClass}`} style={alignStyle}>
            {title}
          </h2>
        )}
        {chip && <span className={styles.chip}>{chip}</span>}
      </div>
    );
  }

  return (
    <div className={`${styles.header} ${styles.headerEdit}`} style={{ justifyContent: justify }}>
      <div className={styles.headingEditWrap} style={{ flex: '1 1 auto' }}>
        <div className={styles.headingToolbar} style={{ alignSelf: justify }} role="toolbar" aria-label="Heading format">
          {SIZES.map((s) => (
            <button
              key={s.key}
              type="button"
              className={`${styles.toolButton} ${size === s.key ? styles.toolActive : ''}`}
              aria-pressed={size === s.key}
              title={`${s.name} heading`}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => onFormatChange(s.key, align)}
            >
              {s.label}
            </button>
          ))}
          <span className={styles.toolDivider} aria-hidden="true" />
          {(['left', 'center', 'right'] as TitleAlign[]).map((a) => (
            <button
              key={a}
              type="button"
              className={`${styles.toolButton} ${align === a ? styles.toolActive : ''}`}
              aria-pressed={align === a}
              title={`Align ${a === 'center' ? 'centre' : a}`}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => onFormatChange(size, a)}
            >
              <AlignIcon align={a} />
            </button>
          ))}
        </div>
        <textarea
          ref={areaRef}
          className={`${styles.heading} ${styles.headingInput} ${sizeClass}`}
          style={alignStyle}
          value={text}
          rows={1}
          placeholder={placeholder}
          aria-label="Web part heading"
          maxLength={255}
          onChange={(e) => {
            // Headings are a single line of text; Enter shouldn't add a line break
            const v = e.target.value.replace(/[\r\n]+/g, ' ');
            setText(v);
            onTitleChange(v);
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') e.preventDefault();
          }}
        />
      </div>
      {chip && <span className={styles.chip}>{chip}</span>}
    </div>
  );
};
