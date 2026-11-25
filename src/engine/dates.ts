// Date serial numbers using the 1900 date system (serial 1 = 1900-01-01).
// For compatibility with the industry-standard 1900 system, serial 60 is the
// fictitious 1900-02-29; serials from 61 onward map to real dates.

const MS_PER_DAY = 86_400_000;
const EPOCH = Date.UTC(1899, 11, 30);

export function serialFromUTC(ms: number): number {
  let s = (ms - EPOCH) / MS_PER_DAY;
  if (s < 61) s -= 1; // before 1900-03-01 the 1900 system is offset by the fictitious leap day
  return s;
}

export function serialFromYMD(y: number, m: number, d: number): number {
  // Date.UTC handles month/day overflow and underflow, matching DATE() semantics.
  const dt = new Date(Date.UTC(2000, 0, 1));
  dt.setUTCFullYear(y, m - 1, d);
  return Math.round(serialFromUTC(dt.getTime()));
}

export function utcFromSerial(serial: number): number {
  let s = serial;
  if (s < 61) s += 1;
  return EPOCH + s * MS_PER_DAY;
}

export interface DateParts {
  y: number;
  m: number; // 1-12
  d: number;
  h: number;
  mi: number;
  s: number;
  ms: number;
  dow: number; // 0 = Sunday
}

export function partsFromSerial(serial: number): DateParts {
  const whole = Math.floor(serial);
  let frac = serial - whole;
  if (whole === 60) {
    // fictitious 1900-02-29
    return { y: 1900, m: 2, d: 29, ...timeParts(frac), dow: 3 };
  }
  const dt = new Date(utcFromSerial(whole));
  // round time to nearest millisecond to avoid 59.9999s artefacts
  let totalMs = Math.round(frac * MS_PER_DAY);
  if (totalMs >= MS_PER_DAY) totalMs = MS_PER_DAY - 1;
  frac = totalMs / MS_PER_DAY;
  return {
    y: dt.getUTCFullYear(),
    m: dt.getUTCMonth() + 1,
    d: dt.getUTCDate(),
    ...timeParts(frac),
    dow: dt.getUTCDay(),
  };
}

function timeParts(frac: number) {
  const totalMs = Math.round(frac * MS_PER_DAY);
  const h = Math.floor(totalMs / 3_600_000);
  const mi = Math.floor((totalMs % 3_600_000) / 60_000);
  const s = Math.floor((totalMs % 60_000) / 1000);
  const ms = totalMs % 1000;
  return { h, mi, s, ms };
}

export function serialNow(now: Date = new Date()): number {
  // Local wall-clock time expressed as a serial.
  const local = Date.UTC(
    now.getFullYear(), now.getMonth(), now.getDate(),
    now.getHours(), now.getMinutes(), now.getSeconds(), now.getMilliseconds(),
  );
  return serialFromUTC(local);
}

export function serialToday(now: Date = new Date()): number {
  return serialFromYMD(now.getFullYear(), now.getMonth() + 1, now.getDate());
}

export function daysInMonth(y: number, m: number): number {
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

export const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
export const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** Parse common date/time text. Returns serial or null. */
export function parseDateText(text: string): { serial: number; kind: 'date' | 'datetime' | 'time'; iso: boolean } | null {
  const t = text.trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t);
  if (m) {
    const y = +m[1], mo = +m[2], d = +m[3];
    if (mo < 1 || mo > 12 || d < 1 || d > daysInMonth(y, mo)) return null;
    let serial = serialFromYMD(y, mo, d);
    if (m[4] !== undefined) {
      const h = +m[4], mi = +m[5], s = m[6] ? +m[6] : 0;
      if (h > 23 || mi > 59 || s > 59) return null;
      serial += (h * 3600 + mi * 60 + s) / 86400;
      return { serial, kind: 'datetime', iso: true };
    }
    return { serial, kind: 'date', iso: true };
  }
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?)?$/.exec(t);
  if (m) {
    const mo = +m[1], d = +m[2];
    let y = +m[3];
    if (m[3].length === 2) y += y < 30 ? 2000 : 1900;
    if (mo < 1 || mo > 12 || d < 1 || d > daysInMonth(y, mo)) return null;
    let serial = serialFromYMD(y, mo, d);
    if (m[4] !== undefined) {
      let h = +m[4];
      const mi = +m[5], s = m[6] ? +m[6] : 0;
      if (m[7]) {
        const pm = m[7].toLowerCase() === 'pm';
        if (h < 1 || h > 12) return null;
        h = (h % 12) + (pm ? 12 : 0);
      }
      if (h > 23 || mi > 59 || s > 59) return null;
      serial += (h * 3600 + mi * 60 + s) / 86400;
      return { serial, kind: 'datetime', iso: false };
    }
    return { serial, kind: 'date', iso: false };
  }
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?$/.exec(t);
  if (m) {
    let h = +m[1];
    const mi = +m[2], s = m[3] ? +m[3] : 0;
    if (m[4]) {
      const pm = m[4].toLowerCase() === 'pm';
      if (h < 1 || h > 12) return null;
      h = (h % 12) + (pm ? 12 : 0);
    }
    if (h > 23 || mi > 59 || s > 59) return null;
    return { serial: (h * 3600 + mi * 60 + s) / 86400, kind: 'time', iso: false };
  }
  // "Jan 5, 2026" / "5 Jan 2026"
  m = /^([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})$/.exec(t) || null;
  let mon: string | undefined, dd: number | undefined, yy: number | undefined;
  if (m) {
    mon = m[1]; dd = +m[2]; yy = +m[3];
  } else {
    const m2 = /^(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})$/.exec(t);
    if (m2) {
      mon = m2[2]; dd = +m2[1]; yy = +m2[3];
    }
  }
  if (mon && dd && yy) {
    const idx = MONTHS.findIndex((x) => x.toLowerCase().startsWith(mon!.toLowerCase().slice(0, 3)) && x.toLowerCase().startsWith(mon!.toLowerCase()));
    if (idx < 0 || dd < 1 || dd > daysInMonth(yy, idx + 1)) return null;
    return { serial: serialFromYMD(yy, idx + 1, dd), kind: 'date', iso: false };
  }
  return null;
}
