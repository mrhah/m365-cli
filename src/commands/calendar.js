import graphClient from '../graph/client.js';
import { outputCalendarList, outputCalendarDetail, outputCalendarResult } from '../utils/output.js';
import { handleError } from '../utils/error.js';

/**
 * Calendar commands
 */

/**
 * List calendar events
 */
export async function listEvents(options) {
  try {
    const { days = 7, top = 50, json = false } = options;
    
    // Calculate date range
    const startDateTime = new Date();
    const endDateTime = new Date();
    endDateTime.setDate(endDateTime.getDate() + days);
    
    const events = await graphClient.calendar.list({
      startDateTime: startDateTime.toISOString(),
      endDateTime: endDateTime.toISOString(),
      top,
    });
    
    outputCalendarList(events, { json, days });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

/**
 * Get calendar event by ID
 */
export async function getEvent(id, options) {
  try {
    const { json = false } = options;
    
    if (!id) {
      throw new Error('Event ID is required');
    }
    
    const event = await graphClient.calendar.get(id);
    
    outputCalendarDetail(event, { json });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

/**
 * Create calendar event
 */
export async function createEvent(title, options) {
  try {
    const {
      start,
      end,
      location = '',
      body = '',
      attendees = [],
      allday = false,
      json = false,
    } = options;
    
    if (!title) {
      throw new Error('Event title is required');
    }
    
    if (!start || !end) {
      throw new Error('Start and end time are required');
    }
    
    // Parse datetime strings
    let startDateTime, endDateTime;
    
    if (allday) {
      // All-day events use date only (YYYY-MM-DD)
      startDateTime = start.includes('T') ? start.split('T')[0] : start;
      endDateTime = end.includes('T') ? end.split('T')[0] : end;
    } else {
      // Regular events use full datetime
      startDateTime = start.includes('T') ? start : `${start}T00:00:00`;
      endDateTime = end.includes('T') ? end : `${end}T00:00:00`;
    }
    
    // Build event object
    const event = {
      subject: title,
      start: {
        dateTime: startDateTime,
        timeZone: 'Asia/Shanghai',
      },
      end: {
        dateTime: endDateTime,
        timeZone: 'Asia/Shanghai',
      },
      isAllDay: allday,
    };
    
    // Add location if provided
    if (location) {
      event.location = {
        displayName: location,
      };
    }
    
    // Add body if provided
    if (body) {
      event.body = {
        contentType: 'HTML',
        content: body,
      };
    }
    
    // Add attendees if provided
    if (attendees && attendees.length > 0) {
      event.attendees = attendees.map(email => ({
        emailAddress: {
          address: email.trim(),
        },
        type: 'required',
      }));
    }
    
    // Create event
    const created = await graphClient.calendar.create(event);
    
    // Build result with properly formatted times
    // Note: Graph API returns dateTime in the specified timeZone, but without zone suffix
    // We need to add the timezone suffix for correct display
    const getTimeWithZone = (timeObj) => {
      const tz = timeObj.timeZone;
      if (tz === 'Asia/Shanghai') {
        return timeObj.dateTime + '+08:00';
      } else if (tz === 'UTC') {
        return timeObj.dateTime + 'Z';
      } else {
        // For other timezones, use as-is (may not display correctly)
        return timeObj.dateTime;
      }
    };
    
    const result = {
      status: 'created',
      id: created.id,
      subject: created.subject,
      start: getTimeWithZone(created.start),
      end: getTimeWithZone(created.end),
    };
    
    outputCalendarResult(result, { json });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

/**
 * Update calendar event
 */
export async function updateEvent(id, options) {
  try {
    const {
      title,
      start,
      end,
      location,
      body,
      json = false,
    } = options;
    
    if (!id) {
      throw new Error('Event ID is required');
    }
    
    // Build update object (only include fields that are provided)
    const updates = {};
    
    if (title) {
      updates.subject = title;
    }
    
    if (start) {
      const startDateTime = start.includes('T') ? start : `${start}T00:00:00`;
      updates.start = {
        dateTime: startDateTime,
        timeZone: 'Asia/Shanghai',
      };
    }
    
    if (end) {
      const endDateTime = end.includes('T') ? end : `${end}T00:00:00`;
      updates.end = {
        dateTime: endDateTime,
        timeZone: 'Asia/Shanghai',
      };
    }
    
    if (location !== undefined) {
      updates.location = {
        displayName: location,
      };
    }
    
    if (body !== undefined) {
      updates.body = {
        contentType: 'HTML',
        content: body,
      };
    }
    
    if (Object.keys(updates).length === 0) {
      throw new Error('No updates provided. Use --title, --start, --end, --location, or --body');
    }
    
    // Update event
    const updated = await graphClient.calendar.update(id, updates);
    
    // Build result with properly formatted times
    // Note: Graph API returns dateTime in the specified timeZone, but without zone suffix
    // We need to add the timezone suffix for correct display
    const getTimeWithZone = (timeObj) => {
      const tz = timeObj.timeZone;
      if (tz === 'Asia/Shanghai') {
        return timeObj.dateTime + '+08:00';
      } else if (tz === 'UTC') {
        return timeObj.dateTime + 'Z';
      } else {
        // For other timezones, use as-is (may not display correctly)
        return timeObj.dateTime;
      }
    };
    
    const result = {
      status: 'updated',
      id: updated.id,
      subject: updated.subject,
      start: getTimeWithZone(updated.start),
      end: getTimeWithZone(updated.end),
    };
    
    outputCalendarResult(result, { json });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

/**
 * Delete calendar event
 */
export async function deleteEvent(id, options) {
  try {
    const { json = false } = options;
    
    if (!id) {
      throw new Error('Event ID is required');
    }
    
    await graphClient.calendar.delete(id);
    
    const result = {
      status: 'deleted',
      id,
    };
    
    outputCalendarResult(result, { json });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

export default {
  list: listEvents,
  get: getEvent,
  create: createEvent,
  update: updateEvent,
  delete: deleteEvent,
};
