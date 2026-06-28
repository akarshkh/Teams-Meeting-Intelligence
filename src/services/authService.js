const { ClientSecretCredential } = require('@azure/identity');
const config = require('../config');
const logger = require('../utils/logger');

class AuthService {
  constructor() {
    this.credential = null;
    this.cachedToken = null;
    this.tokenExpiresAt = 0;
  }

  /**
   * Initializes the ClientSecretCredential.
   */
  getCredential() {
    if (!this.credential) {
      if (!config.tenantId || !config.clientId || !config.clientSecret) {
        throw new Error('Missing client credentials in config. Please set TENANT_ID, CLIENT_ID, and CLIENT_SECRET.');
      }
      this.credential = new ClientSecretCredential(
        config.tenantId,
        config.clientId,
        config.clientSecret
      );
    }
    return this.credential;
  }

  /**
   * Retrieves an access token for Microsoft Graph.
   * Leverages caching to minimize network calls.
   * 
   * @returns {Promise<string>} The access token.
   */
  async getAccessToken() {
    const now = Date.now();
    
    // Return cached token if it's still valid (with a 5-minute buffer)
    if (this.cachedToken && this.tokenExpiresAt > now + 5 * 60 * 1000) {
      return this.cachedToken;
    }

    logger.info('Acquiring new Microsoft Graph access token...');
    try {
      const credential = this.getCredential();
      const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
      
      this.cachedToken = tokenResponse.token;
      this.tokenExpiresAt = tokenResponse.expiresOnTimestamp;
      
      logger.info('Access token acquired successfully.');
      return this.cachedToken;
    } catch (error) {
      logger.error('Failed to acquire access token.', error);
      throw error;
    }
  }
}

module.exports = new AuthService();
