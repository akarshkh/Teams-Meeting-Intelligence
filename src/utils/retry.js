const logger = require('./logger');

/**
 * Executes a function and retries it with exponential backoff and jitter.
 * 
 * @param {Function} fn The async function to execute.
 * @param {Object} options Retry configuration.
 * @param {number} options.retries Maximum number of retry attempts.
 * @param {number} options.factor The exponential factor.
 * @param {number} options.minTimeout Minimum time to wait in milliseconds.
 * @param {number} options.maxTimeout Maximum time to wait in milliseconds.
 * @param {string} options.operationName Name of the operation for logging.
 */
async function retry(fn, { retries = 5, factor = 2, minTimeout = 1000, maxTimeout = 10000, operationName = 'Operation' } = {}) {
  let attempt = 0;
  
  while (true) {
    try {
      attempt++;
      return await fn();
    } catch (error) {
      if (attempt > retries) {
        logger.error(`${operationName} failed after ${retries} retries. Re-throwing error.`, error);
        throw error;
      }
      
      // Calculate exponential backoff
      let delay = minTimeout * Math.pow(factor, attempt - 1);
      // Clamp to max timeout
      delay = Math.min(delay, maxTimeout);
      // Add random jitter (-20% to +20%)
      const jitter = (Math.random() * 0.4 - 0.2) * delay;
      const sleepTime = Math.max(0, Math.round(delay + jitter));
      
      logger.warn(`${operationName} attempt ${attempt} failed. Retrying in ${sleepTime}ms. Error: ${error.message}`);
      
      await new Promise(resolve => setTimeout(resolve, sleepTime));
    }
  }
}

module.exports = retry;
