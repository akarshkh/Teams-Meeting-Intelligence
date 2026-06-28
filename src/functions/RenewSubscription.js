const { app } = require('@azure/functions');
const graphService = require('../services/graphService');
const config = require('../config');
const logger = require('../utils/logger');

app.timer('RenewSubscription', {
  schedule: '0 0 */12 * * *', // Runs every 12 hours (standard cron: second, minute, hour, day, month, dayOfWeek)
  handler: async (myTimer, context) => {
    logger.info('Timer Trigger [RenewSubscription] activated.');

    if (!config.webhookUrl) {
      logger.warn('WEBHOOK_URL is not configured. Skipping subscription renewal check.');
      return;
    }

    try {
      // 1. Fetch all active subscriptions from Microsoft Graph
      const subscriptions = await graphService.listSubscriptions();
      
      // Target resource
      const targetResource = 'communications/onlineMeetings/getAllTranscripts';

      // 2. Look for an existing subscription matching our webhook URL and resource
      const existingSub = subscriptions.find(sub => 
        sub.resource === targetResource && 
        sub.notificationUrl === config.webhookUrl
      );

      if (existingSub) {
        logger.info(`Found existing subscription (ID: ${existingSub.id}) expiring at: ${existingSub.expirationDateTime}`);
        
        // Calculate new expiration (2 days from now, well within the 3 days limit)
        const newExpiration = new Date(Date.now() + 2 * 24 * 60 * 60 * 1000).toISOString();
        
        // Renew the subscription
        await graphService.renewSubscription(existingSub.id, newExpiration);
        logger.info(`Subscription ${existingSub.id} renewed successfully. New Expiration: ${newExpiration}`);
      } else {
        logger.warn('No active subscription found matching this webhook configuration. Creating a new subscription...');
        
        // Create a new subscription
        // Default to a 2-day expiration to ensure it is robust, providing a lifecycle URL
        const newExpiration = new Date(Date.now() + 2 * 24 * 60 * 60 * 1000).toISOString();
        const newSub = await graphService.createSubscription({
          resource: targetResource,
          expirationDateTime: newExpiration
        });
        
        logger.info(`New subscription created successfully. ID: ${newSub.id}, Expires: ${newSub.expirationDateTime}`);
      }
    } catch (error) {
      logger.error('Error during automatic subscription renewal check.', error);
    }
  }
});
