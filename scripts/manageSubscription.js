const fs = require('fs');
const path = require('path');


// Load settings from local.settings.json if process.env variables are empty (e.g. running outside func start)
function loadLocalSettings() {
  try {
    const settingsPath = path.join(__dirname, '../local.settings.json');
    if (fs.existsSync(settingsPath)) {
      const content = fs.readFileSync(settingsPath, 'utf8');
      const parsed = JSON.parse(content);
      if (parsed.Values) {
        for (const [key, value] of Object.entries(parsed.Values)) {
          if (!process.env[key]) {
            process.env[key] = value;
          }
        }
      }
    }
  } catch (err) {
    // Suppress warning if file is missing, we will fallback to standard process.env
  }
}


loadLocalSettings();

const graphService = require('../src/services/graphService');
const logger = require('../src/utils/logger');
// Re-initialize config requirements
const config = require('../src/config');

function printUsage() {
  console.log(`
Meridian Teams Meeting Intelligence - Subscription Manager CLI

Usage:
  npm run manage-sub <command> [arguments]

Commands:
  list                              List all active subscriptions for the app.
  create [resource] [expiry]        Create a new subscription.
                                    - resource defaults to: communications/onlineMeetings/getAllTranscripts
                                    - expiry defaults to: 55 minutes from now (without lifecycle URL)
                                      or 2 days from now (if LIFECYCLE_NOTIFICATION_URL is set)
  renew <subscriptionId> [expiry]   Renew a subscription.
                                    - expiry defaults to 2 days from now
  delete <subscriptionId>           Delete a subscription.

Examples:
  npm run manage-sub list
  npm run manage-sub create
  npm run manage-sub delete 8fb8a2cd-ec05-4c07-ba21-140237fb57d4
`);
}

async function run() {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    printUsage();
    process.exit(0);
  }

  const command = args[0].toLowerCase();

  try {
    switch (command) {
      case 'list': {
        const subs = await graphService.listSubscriptions();
        console.log('\nActive Subscriptions:');
        if (subs.length === 0) {
          console.log('No active subscriptions found.');
        } else {
          console.table(subs.map(s => ({
            id: s.id,
            resource: s.resource,
            expiration: s.expirationDateTime,
            url: s.notificationUrl
          })));
        }
        break;
      }

      case 'create': {
        const resource = args[1] || 'communications/onlineMeetings/getAllTranscripts';
        
        // If lifecycle URL is set, default to 2 days, else 55 minutes (Graph max without lifecycle URL is 1 hr)
        let defaultExpiry;
        if (config.lifecycleNotificationUrl) {
          defaultExpiry = new Date(Date.now() + 2 * 24 * 60 * 60 * 1000).toISOString(); // 2 days
        } else {
          defaultExpiry = new Date(Date.now() + 55 * 60 * 1000).toISOString(); // 55 mins
        }
        
        const expiry = args[2] || defaultExpiry;
        const sub = await graphService.createSubscription({ resource, expirationDateTime: expiry });
        console.log('\nSubscription Created:');
        console.log(JSON.stringify(sub, null, 2));
        break;
      }

      case 'renew': {
        const subId = args[1];
        if (!subId) {
          console.error('Error: Please provide a subscription ID.');
          printUsage();
          process.exit(1);
        }
        const defaultExpiry = new Date(Date.now() + 2 * 24 * 60 * 60 * 1000).toISOString(); // 2 days
        const expiry = args[2] || defaultExpiry;
        const sub = await graphService.renewSubscription(subId, expiry);
        console.log('\nSubscription Renewed:');
        console.log(JSON.stringify(sub, null, 2));
        break;
      }

      case 'delete': {
        const subId = args[1];
        if (!subId) {
          console.error('Error: Please provide a subscription ID.');
          printUsage();
          process.exit(1);
        }
        await graphService.deleteSubscription(subId);
        console.log(`\nSubscription ${subId} successfully deleted.`);
        break;
      }

      default:
        console.error(`Unknown command: ${command}`);
        printUsage();
        process.exit(1);
    }
  } catch (err) {
    console.error('\nCLI Operation Failed:');
    if (err.response && err.response.data) {
      console.error(JSON.stringify(err.response.data, null, 2));
    } else {
      console.error(err.message);
    }
    process.exit(1);
  }
}

run();
