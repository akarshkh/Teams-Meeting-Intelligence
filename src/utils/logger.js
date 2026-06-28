/**
 * Structured Logging Utility for Meridian Teams Meeting Intelligence Webhook
 */

function formatLog(level, message, context = {}) {
  const logEntry = {
    timestamp: new Date().toISOString(),
    level,
    message,
    ...context
  };
  return JSON.stringify(logEntry);
}

const logger = {
  info: (message, context) => {
    console.log(formatLog('INFO', message, context));
  },
  warn: (message, context) => {
    console.warn(formatLog('WARN', message, context));
  },
  error: (message, error, context = {}) => {
    const errorDetails = error ? { 
      errorName: error.name,
      errorMessage: error.message,
      errorStack: error.stack 
    } : {};
    console.error(formatLog('ERROR', message, { ...context, ...errorDetails }));
  },
  
  /**
   * Log an audit record when a transcript is successfully processed and sent to Power Automate.
   */
  auditSuccess: ({ meetingId, transcriptId, subject, organizer, transcriptLength, processingTimeMs }) => {
    console.log(formatLog('AUDIT_SUCCESS', 'Transcript processed successfully and sent to Power Automate.', {
      meetingId,
      transcriptId,
      meetingSubject: subject,
      organizer,
      transcriptLength,
      processingTimeMs: `${processingTimeMs}ms`
    }));
  },

  /**
   * Log a dead letter event when transcript processing fails completely after all retries.
   */
  deadLetter: (message, rawNotification, error, context = {}) => {
    const errorDetails = error ? { 
      errorName: error.name,
      errorMessage: error.message,
      errorStack: error.stack 
    } : {};
    console.error(formatLog('DEAD_LETTER', message, {
      ...context,
      rawNotification,
      ...errorDetails
    }));
  }
};

module.exports = logger;
