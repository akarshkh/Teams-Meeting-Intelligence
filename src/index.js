const { app } = require('@azure/functions');

app.setup({
    enableHttpStream: true,
});

require('./functions/GraphWebhook');
require('./functions/Health');
require('./functions/RenewSubscription');