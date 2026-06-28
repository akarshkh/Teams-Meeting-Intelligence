const axios = require('axios');
const authService = require('./authService');
const config = require('../config');
const logger = require('../utils/logger');
const retry = require('../utils/retry');

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

class GraphService {
  /**
   * Helper to get request headers with auth token.
   */
  async _getHeaders(extraHeaders = {}) {
    const token = await authService.getAccessToken();
    return {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...extraHeaders
    };
  }

  /**
   * Creates a subscription for online meeting transcript events.
   * Monitors the tenant-wide communications/onlineMeetings/getAllTranscripts resource.
   */
  async createSubscription({ resource, expirationDateTime }) {
    const targetResource = resource || 'communications/onlineMeetings/getAllTranscripts';
    
    // Default expiration to 1 hour (60 minutes) if not specified to prevent lifecycle URL issues,
    // or up to 3 days (4320 mins) if lifecycle notification URL is provided.
    const defaultExpiration = new Date(Date.now() + 55 * 60 * 1000).toISOString(); // 55 mins
    const expiry = expirationDateTime || defaultExpiration;

    const subscriptionPayload = {
      changeType: 'created',
      notificationUrl: config.webhookUrl,
      resource: targetResource,
      expirationDateTime: expiry,
      clientState: config.clientState
    };

    // If expiration is more than 1 hour in the future, Microsoft Graph requires a lifecycle URL
    const diffMs = new Date(expiry) - Date.now();
    if (diffMs > 60 * 60 * 1000) {
      if (!config.lifecycleNotificationUrl) {
        throw new Error('LIFECYCLE_NOTIFICATION_URL must be configured if subscription expiration is > 1 hour.');
      }
      subscriptionPayload.lifecycleNotificationUrl = config.lifecycleNotificationUrl;
    }

    logger.info(`Creating Graph subscription on resource: ${targetResource} expiring at ${expiry}`);
    
    return retry(async () => {
      const headers = await this._getHeaders();
      try {
    const response = await axios.post(
        `${GRAPH_BASE_URL}/subscriptions`,
        subscriptionPayload,
        { headers }
    );

    logger.info(`Subscription created successfully. ID: ${response.data.id}`);
    return response.data;

} catch (err) {
    console.error("GRAPH ERROR:");
    console.error(JSON.stringify(err.response?.data, null, 2));
    throw err;
}
      logger.info(`Subscription created successfully. ID: ${response.data.id}`);
      return response.data;
    }, { operationName: 'Create Graph Subscription' });
  }

  /**
   * Lists all active subscriptions for the application.
   */
  async listSubscriptions() {
    logger.info('Listing all active subscriptions...');
    return retry(async () => {
      const headers = await this._getHeaders();
      const response = await axios.get(`${GRAPH_BASE_URL}/subscriptions`, { headers });
      return response.data.value || [];
    }, { operationName: 'List Graph Subscriptions' });
  }

  /**
   * Retrieves an active subscription from Microsoft Graph.
   */
  async getSubscription(subscriptionId) {
    logger.info(`Fetching subscription details for: ${subscriptionId}`);
    return retry(async () => {
      const headers = await this._getHeaders();
      const response = await axios.get(`${GRAPH_BASE_URL}/subscriptions/${subscriptionId}`, { headers });
      return response.data;
    }, { operationName: 'Get Graph Subscription' });
  }

  /**
   * Renews an existing subscription by extending its expiration.
   */
  async renewSubscription(subscriptionId, expirationDateTime) {
    // Default to extend by 2 days (max is 3 days)
    const expiry = expirationDateTime || new Date(Date.now() + 2 * 24 * 60 * 60 * 1000).toISOString();
    
    logger.info(`Renewing subscription: ${subscriptionId} with new expiration: ${expiry}`);
    
    const payload = {
      expirationDateTime: expiry
    };

    // Include lifecycle URL if renewing for more than 1 hour
    if (config.lifecycleNotificationUrl) {
      payload.lifecycleNotificationUrl = config.lifecycleNotificationUrl;
    }

    return retry(async () => {
      const headers = await this._getHeaders();
      const response = await axios.patch(`${GRAPH_BASE_URL}/subscriptions/${subscriptionId}`, payload, { headers });
      logger.info(`Subscription ${subscriptionId} renewed successfully.`);
      return response.data;
    }, { operationName: 'Renew Graph Subscription' });
  }

  /**
   * Deletes a subscription.
   */
  async deleteSubscription(subscriptionId) {
    logger.info(`Deleting subscription: ${subscriptionId}`);
    return retry(async () => {
      const headers = await this._getHeaders();
      await axios.delete(`${GRAPH_BASE_URL}/subscriptions/${subscriptionId}`, { headers });
      logger.info(`Subscription ${subscriptionId} deleted successfully.`);
      return true;
    }, { operationName: 'Delete Graph Subscription' });
  }

  /**
   * Retrieves online meeting metadata (to resolve Subject and Organizer/User ID).
   */
  async getMeeting(userId, meetingId) {
    logger.info(`Fetching meeting details for User ID: ${userId}, Meeting ID: ${meetingId}`);
    return retry(async () => {
      const headers = await this._getHeaders();
      try {
        const response = await axios.get(`${GRAPH_BASE_URL}/users/${userId}/onlineMeetings/${meetingId}`, { headers });
        return response.data;
      } catch (err) {
        console.log('========== GRAPH GET MEETING ERROR ==========');
        console.log('Status:', err.response?.status);
        console.log('Headers:', JSON.stringify(err.response?.headers, null, 2));
        console.log('Body:', JSON.stringify(err.response?.data, null, 2));
        console.log('============================================');
        throw err;
      }
    }, { operationName: 'Get Online Meeting Metadata' });
  }

  /**
   * Downloads the raw WebVTT (.vtt) transcript content.
   */
  async getTranscriptContent(userId, meetingId, transcriptId) {
    logger.info(
      `Retrieving transcript content. Meeting ID: ${meetingId}, Transcript ID: ${transcriptId}`
    );

    return retry(async () => {
      const headers = await this._getHeaders({
        Accept: 'text/vtt'
      });

      try {
        const response = await axios.get(
          `${GRAPH_BASE_URL}/users/${userId}/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content`,
          {
            headers,
            responseType: 'text'
          }
        );

        return response.data;
      } catch (err) {
        console.log('========== GRAPH GET TRANSCRIPT ERROR ==========');
        console.log('Status:', err.response?.status);
        console.log('Body:', JSON.stringify(err.response?.data, null, 2));
        console.log('===============================================');
        throw err;
      }
    }, {
      operationName: 'Get Transcript Content'
    });
  }
}

module.exports = new GraphService();
