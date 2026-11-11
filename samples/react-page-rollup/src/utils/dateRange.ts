import { DateRange } from '../models/IPageRollupProps';

export interface IDateWindow {
  start?: Date;
  end?: Date;
}

/** Works out the start/end of the chosen date window, in the viewer's local time. */
export function getDateWindow(range: DateRange, lastDays: number, now: Date = new Date()): IDateWindow {
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const quarterStartMonth = Math.floor(now.getMonth() / 3) * 3;

  switch (range) {
    case 'lastDays': {
      const days = Math.max(1, Math.round(lastDays || 30));
      const start = new Date(today);
      start.setDate(start.getDate() - days);
      return { start };
    }
    case 'thisQuarter':
      return {
        start: new Date(now.getFullYear(), quarterStartMonth, 1),
        end: new Date(now.getFullYear(), quarterStartMonth + 3, 1)
      };
    case 'lastQuarter':
      return {
        start: new Date(now.getFullYear(), quarterStartMonth - 3, 1),
        end: new Date(now.getFullYear(), quarterStartMonth, 1)
      };
    case 'thisYear':
      return {
        start: new Date(now.getFullYear(), 0, 1),
        end: new Date(now.getFullYear() + 1, 0, 1)
      };
    default:
      return {};
  }
}

/** Short human description, e.g. "Q4 2026" or "Last 30 days". */
export function describeDateWindow(range: DateRange, lastDays: number, now: Date = new Date()): string {
  const q = Math.floor(now.getMonth() / 3) + 1;
  switch (range) {
    case 'lastDays':
      return `Last ${Math.max(1, Math.round(lastDays || 30))} days`;
    case 'thisQuarter':
      return `Q${q} ${now.getFullYear()}`;
    case 'lastQuarter':
      return q === 1 ? `Q4 ${now.getFullYear() - 1}` : `Q${q - 1} ${now.getFullYear()}`;
    case 'thisYear':
      return `${now.getFullYear()}`;
    default:
      return '';
  }
}
