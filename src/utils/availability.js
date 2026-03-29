const STATUS_MAP = {
  '0': 'free',
  '1': 'tentative',
  '2': 'busy',
  '3': 'oof',
  '4': 'workingElsewhere',
};

function pad(value) {
  return String(value).padStart(2, '0');
}

function isLeapYear(year) {
  return (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
}

function daysInMonth(year, month) {
  if (month === 2) {
    return isLeapYear(year) ? 29 : 28;
  }

  if ([4, 6, 9, 11].includes(month)) {
    return 30;
  }

  return 31;
}

function parseDateTime(dateTimeStr) {
  if (typeof dateTimeStr !== 'string') {
    return null;
  }

  const base = dateTimeStr.slice(0, 19);
  const match = base.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) {
    return null;
  }

  return {
    year: Number.parseInt(match[1], 10),
    month: Number.parseInt(match[2], 10),
    day: Number.parseInt(match[3], 10),
    hour: Number.parseInt(match[4], 10),
    minute: Number.parseInt(match[5], 10),
    second: Number.parseInt(match[6] || '0', 10),
  };
}

function formatDateTime(parts) {
  return `${parts.year}-${pad(parts.month)}-${pad(parts.day)}T${pad(parts.hour)}:${pad(parts.minute)}:${pad(parts.second)}`;
}

export function addMinutes(dateTimeStr, minutesToAdd) {
  const parsed = parseDateTime(dateTimeStr);
  if (!parsed) {
    return dateTimeStr;
  }

  let year = parsed.year;
  let month = parsed.month;
  let day = parsed.day;
  let hour = parsed.hour;
  let minute = parsed.minute + minutesToAdd;
  const second = parsed.second;

  while (minute >= 60) {
    minute -= 60;
    hour += 1;
  }

  while (hour >= 24) {
    hour -= 24;
    day += 1;
    const dim = daysInMonth(year, month);
    if (day > dim) {
      day = 1;
      month += 1;

      if (month > 12) {
        month = 1;
        year += 1;
      }
    }
  }

  while (minute < 0) {
    minute += 60;
    hour -= 1;
  }

  while (hour < 0) {
    hour += 24;
    day -= 1;

    if (day < 1) {
      month -= 1;
      if (month < 1) {
        month = 12;
        year -= 1;
      }
      day = daysInMonth(year, month);
    }
  }

  return formatDateTime({ year, month, day, hour, minute, second });
}

function daysBeforeYear(year) {
  const y = year - 1;
  return (y * 365)
    + Math.floor(y / 4)
    - Math.floor(y / 100)
    + Math.floor(y / 400);
}

function dayOfYear(year, month, day) {
  const monthDays = [31, isLeapYear(year) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  let total = 0;
  for (let i = 0; i < month - 1; i += 1) {
    total += monthDays[i];
  }
  return total + day;
}

function toAbsoluteMinutes(dateTimeStr) {
  const parsed = parseDateTime(dateTimeStr);
  if (!parsed) {
    return null;
  }

  const days = daysBeforeYear(parsed.year) + dayOfYear(parsed.year, parsed.month, parsed.day);
  return (days * 24 * 60) + (parsed.hour * 60) + parsed.minute;
}

function diffMinutes(start, end) {
  const startAbs = toAbsoluteMinutes(start);
  const endAbs = toAbsoluteMinutes(end);

  if (startAbs === null || endAbs === null) {
    return 0;
  }

  return Math.max(0, endAbs - startAbs);
}

export function mapStatus(char) {
  return STATUS_MAP[char] || 'unknown';
}

export function expandSlots(view, startDateTime, intervalMinutes) {
  if (!view) {
    return [];
  }

  const safeInterval = Number.isFinite(intervalMinutes) && intervalMinutes > 0 ? intervalMinutes : 30;

  return Array.from(view).map((char, index) => {
    const start = addMinutes(startDateTime, index * safeInterval);
    const end = addMinutes(start, safeInterval);

    return {
      start,
      end,
      status: mapStatus(char),
    };
  });
}

export function mergeSegments(slots) {
  if (!Array.isArray(slots) || slots.length === 0) {
    return [];
  }

  const segments = [];
  let current = {
    start: slots[0].start,
    end: slots[0].end,
    status: slots[0].status,
  };

  for (let i = 1; i < slots.length; i += 1) {
    const slot = slots[i];

    if (slot.status === current.status && slot.start === current.end) {
      current.end = slot.end;
      continue;
    }

    segments.push({
      ...current,
      durationMinutes: diffMinutes(current.start, current.end),
    });

    current = {
      start: slot.start,
      end: slot.end,
      status: slot.status,
    };
  }

  segments.push({
    ...current,
    durationMinutes: diffMinutes(current.start, current.end),
  });

  return segments;
}

export function filterFreeSegments(segments) {
  if (!Array.isArray(segments) || segments.length === 0) {
    return [];
  }

  return segments
    .filter(segment => segment.status === 'free')
    .map(({ start, end, durationMinutes }) => ({ start, end, durationMinutes }));
}

export function enrichScheduleResult(scheduleResult, queryContext) {
  const {
    startDateTime,
    endDateTime,
    timeZone,
    intervalMinutes,
  } = queryContext || {};

  const slots = expandSlots(scheduleResult?.availabilityView || '', startDateTime, intervalMinutes);
  const segments = mergeSegments(slots);
  const freeSegments = filterFreeSegments(segments);

  return {
    ...scheduleResult,
    startDateTime,
    endDateTime,
    timeZone,
    intervalMinutes,
    slots,
    segments,
    freeSegments,
  };
}

export default {
  mapStatus,
  addMinutes,
  expandSlots,
  mergeSegments,
  filterFreeSegments,
  enrichScheduleResult,
};
