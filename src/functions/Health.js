const { app } = require('@azure/functions');

app.http('Health', {
  methods: ['GET'],
  route: 'health',
  authLevel: 'anonymous',
  handler: async (request, context) => {
    return {
      status: 200,
      headers: {
        'Content-Type': 'application/json'
      },
      jsonBody: {
        status: "Healthy"
      }
    };
  }
});
