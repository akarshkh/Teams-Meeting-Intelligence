const config = {
  tenantId: process.env.TENANT_ID,
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
  powerAutomateUrl: process.env.POWER_AUTOMATE_URL,
  clientState: process.env.CLIENT_STATE,
  webhookUrl: process.env.WEBHOOK_URL,
  lifecycleNotificationUrl: process.env.LIFECYCLE_NOTIFICATION_URL || process.env.WEBHOOK_URL,
  azureWebJobsStorage: process.env.AzureWebJobsStorage
};

// Validate that required variables are present
const requiredKeys = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET', 'POWER_AUTOMATE_URL', 'CLIENT_STATE'];
const missingKeys = [];

for (const key of requiredKeys) {
  if (!process.env[key]) {
    missingKeys.push(key);
  }
}

if (missingKeys.length > 0) {
  console.warn(`[WARNING] Missing required environment variables: ${missingKeys.join(', ')}. App may not function correctly.`);
}

module.exports = config;
