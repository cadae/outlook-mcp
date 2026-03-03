import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';

// List calendar events
export async function listEventsTool(authManager, args) {
  const { startDateTime, endDateTime, limit = 10, calendar } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Use calendarView endpoint which properly handles timezones and recurring events
    // This is more reliable than /events with filter
    const endpoint = calendar 
      ? `/me/calendars/${calendar}/calendarView` 
      : '/me/calendarView';
    
    const options = {
      select: 'id,subject,start,end,location,attendees,bodyPreview,organizer,isAllDay,showAs,importance,sensitivity,categories,webLink',
      top: limit,
      orderby: 'start/dateTime',
    };

    // calendarView requires startDateTime and endDateTime as query parameters
    if (startDateTime && endDateTime) {
      options.startDateTime = startDateTime;
      options.endDateTime = endDateTime;
    } else {
      // Default to today if no dates specified
      const today = new Date();
      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);
      options.startDateTime = today.toISOString();
      options.endDateTime = tomorrow.toISOString();
    }

    const result = await graphApiClient.makeRequest(endpoint, options);

    const events = result.value.map(event => ({
      id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      location: event.location?.displayName || 'No location',
      attendees: event.attendees?.map(a => a.emailAddress?.address) || [],
      preview: event.bodyPreview,
      organizer: event.organizer?.emailAddress?.address || 'Unknown',
      isAllDay: event.isAllDay,
      webLink: event.webLink,
    }));

    return createSafeResponse({ events, count: events.length });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list events');
  }
}