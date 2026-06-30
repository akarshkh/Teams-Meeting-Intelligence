const { app } = require('@azure/functions');
const axios = require('axios');
const config = require('../config');
const logger = require('../utils/logger');
const retry = require('../utils/retry');
const graphService = require('../services/graphService');

// In-memory cache for duplicate detection (holds transcript IDs that are currently being processed or have been processed)
// To prevent memory growth, we limit it to 5000 entries and clean up periodically.
const processedTranscripts = new Set();
const MAX_CACHE_SIZE = 5000;

function isDuplicate(transcriptId) {
  if (processedTranscripts.has(transcriptId)) {
    return true;
  }
  // Add to cache
  processedTranscripts.add(transcriptId);
  // Housekeeping: remove oldest entries if size exceeded
  if (processedTranscripts.size > MAX_CACHE_SIZE) {
    const iterator = processedTranscripts.values();
    // Remove the first/oldest 100 entries
    for (let i = 0; i < 100; i++) {
      processedTranscripts.delete(iterator.next().value);
    }
  }
  return false;
}

app.http('GraphWebhook', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const startTime = Date.now();

    // 1. Handle Microsoft Graph validation BEFORE reading the request body
    const validationToken = request.query.get("validationToken");

    if (validationToken) {
      logger.info("Microsoft Graph validation request received.");

      return {
        status: 200,
        headers: {
          "Content-Type": "text/plain"
        },
        body: validationToken
      };
    }

    // 2. Handle POST change notifications from Microsoft Graph
    try {
      let requestBody;

      // Parse JSON only after validation
      try {
        requestBody = await request.json();
        logger.info("GRAPH PAYLOAD:");
        logger.info(JSON.stringify(requestBody, null, 2));
        logger.info("CLIENT STATE RECEIVED: " + JSON.stringify(requestBody.value?.[0]?.clientState));
        logger.info("RESOURCE RECEIVED: " + JSON.stringify(requestBody.value?.[0]?.resource));
      } catch (err) {
        // If no JSON body exists, return HTTP 200 with a simple JSON response instead of throwing
        logger.warn("No JSON body received.");
        return {
          status: 200,
          jsonBody: {
            message: "No JSON body."
          }
        };
      }

      if (!requestBody) {
        logger.warn("Request body is empty.");
        return {
          status: 200,
          jsonBody: {
            message: "No JSON body."
          }
        };
      }

      logger.info("Incoming Graph Webhook notification received.");

      // Check if it's a lifecycle notification or standard change notification
      if (!requestBody.value || !Array.isArray(requestBody.value)) {
        // Preserve lifecycle notification handling
        if (requestBody.lifecycleEvent) {
          logger.warn(`Received Microsoft Graph lifecycle notification: ${requestBody.lifecycleEvent}`);
          return {
            status: 202,
            jsonBody: { message: "Lifecycle notification acknowledged." }
          };
        }
        
        logger.warn('Received request with unexpected body format.', { requestBody });
        return {
          status: 400,
          jsonBody: { error: "Invalid request payload format." }
        };
      }

      // Process notifications in parallel
      const processPromises = requestBody.value.map(async (notification) => {
        const itemStartTime = Date.now();

        // A. Validate Client State - Preserve clientState validation
        if (notification.clientState !== config.clientState) {
          logger.warn('ClientState mismatch. Rejecting notification.', {
            received: notification.clientState,
            expected: config.clientState ? '***' : 'undefined'
          });
          return;
        }

        // B. Parse Meeting ID and Transcript ID from the resource URL
        const resource = notification.resource;
       const resourceRegex =
            /users\('([^']+)'\)\/onlineMeetings\('([^']+)'\)\/transcripts\('([^']+)'\)/i;
        const match = resource.match(resourceRegex);

        if (!match) {
          logger.warn(`Could not parse meetingId and transcriptId from resource: ${resource}`);
          return;
        }

        const userId = match[1];
        const meetingId = match[2];
        const transcriptId = match[3];

        // C. Duplicate Detection - Preserve duplicate detection
        if (isDuplicate(transcriptId)) {
          logger.info(`Duplicate transcript notification detected. Skipping. ID: ${transcriptId}`);
          return;
        }

        logger.info(`Processing transcript event. User ID: ${userId}, Meeting ID: ${meetingId}, Transcript ID: ${transcriptId}`);

        try {
          // D. Fetch Meeting Details - Preserve Graph API calls
          // Fetch meeting details from Microsoft Graph
          logger.info(`User ID: ${userId}`);
          logger.info(`Meeting ID: ${meetingId}`);
          logger.info(`Transcript ID: ${transcriptId}`);
          logger.info(`Fetching meeting details for ID: ${meetingId}`);
          const meetingDetails = await graphService.getMeeting(userId, meetingId);
          logger.info("===== GRAPH MEETING DETAILS =====");
          logger.info(JSON.stringify(meetingDetails, null, 2));
          logger.info("================================");
          const meetingSubject = meetingDetails.subject || "Microsoft Teams Meeting";
          const organizerName =
            meetingDetails.participants?.organizer?.identity?.displayName ||
            meetingDetails.participants?.organizer?.user?.displayName ||
            "Unknown Organizer";
          const participants =
            meetingDetails.participants?.attendees
              ?.map((attendee) =>
                attendee.identity?.displayName ||
                attendee.upn ||
                attendee.identity?.user?.displayName
              )
              .filter(Boolean)
              .join(", ") || "Not Available";
          const participantCount = meetingDetails.participants?.attendees?.length || 0;
          const meetingStartTime = meetingDetails.startDateTime || "";
          const meetingEndTime = meetingDetails.endDateTime || "";
          let meetingDuration = "Not Available";
          if (meetingStartTime && meetingEndTime) {
            const start = new Date(meetingStartTime);
            const end = new Date(meetingEndTime);
            const durationMinutes = Math.round((end.getTime() - start.getTime()) / 60000);
            const hours = Math.floor(durationMinutes / 60);
            const minutes = durationMinutes % 60;
            meetingDuration = hours > 0 ? `${hours} hr ${minutes} min` : `${minutes} min`;
          }

          logger.info(`Organizer: ${organizerName}`);
          logger.info(`Participants: ${participants}`);
          logger.info(`Participant Count: ${participantCount}`);
          logger.info(`Meeting Start: ${meetingStartTime}`);
          logger.info(`Meeting End: ${meetingEndTime}`);
          logger.info(`Meeting Duration: ${meetingDuration}`);

          // E. Fetch Transcript VTT Content - Preserve Graph API calls
          logger.info(`Fetching transcript content for ID: ${transcriptId}`);
          const vttContent = await graphService.getTranscriptContent(
    userId,
    meetingId,
    transcriptId
);
          const transcriptLength = vttContent.length;

          // F. Forward to Power Automate via HTTP POST - Preserve Power Automate integration & retry logic
          logger.info('Forwarding transcript and metadata to Power Automate...');
          logger.info(`Power Automate URL: ${config.powerAutomateUrl}`);
          logger.info("Calling Power Automate...");
          await retry(async () => {
            await axios.post(config.powerAutomateUrl, {
              meetingId,
              transcriptId,
              meetingSubject,
              organizerName,
              participants,
              participantCount,
              meetingStartTime,
              meetingEndTime,
              meetingDuration,
              transcriptLength,
              transcriptContent: vttContent,
              processedAt: new Date().toISOString()
            }, {
              headers: { 'Content-Type': 'application/json' },
              timeout: 15000 // 15 seconds timeout
            });
          }, {
            retries: 5,
            factor: 2,
            minTimeout: 1000,
            maxTimeout: 10000,
            operationName: 'Forward to Power Automate'
          });

          // G. Log successful processing audit trail - Preserve audit logging
          const duration = Date.now() - itemStartTime;
          logger.auditSuccess({
            meetingId,
            transcriptId,
            subject: meetingSubject,
            organizer: organizerName,
            transcriptLength,
            processingTimeMs: duration
          });

        } catch (itemError) {
          // H. Dead Letter Logging for individual event failure - Preserve dead-letter logging
          logger.deadLetter(
            `Failed to process transcript event after retries. Meeting ID: ${meetingId}, Transcript ID: ${transcriptId}`,
            notification,
            itemError,
            { meetingId, transcriptId }
          );
        }
      });

      // Wait for all notifications in the batch to complete processing
      await Promise.all(processPromises);

      const totalDuration = Date.now() - startTime;
      logger.info(`Completed processing batch of notifications in ${totalDuration}ms.`);

      return {
        status: 202,
        jsonBody: { message: "Notifications processed." }
      };

    } catch (err) {
      // Ensure there is only ONE outer try { ... } catch
      logger.error('Error processing notification batch.', err);
      return {
        status: 500,
        jsonBody: { error: "An internal error occurred." }
      };
    }
  }
});