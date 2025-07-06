const axios = require('axios');

exports.handler = async function (event, context) {
  const params = event.queryStringParameters;
  const code = params.code;

  if (!code) {
    return { statusCode: 400, body: 'No code provided' };
  }

  try {
    const tokenResponse = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', null, {
      params: {
        client_id: process.env.CLIENT_ID,
        scope: 'offline_access https://graph.microsoft.com/Mail.Read',
        code: code,
        redirect_uri: process.env.REDIRECT_URI,
        grant_type: 'authorization_code',
        client_secret: process.env.CLIENT_SECRET,
      },
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });

    const { refresh_token, access_token } = tokenResponse.data;

    const successPage = `
      <h1>Authentication Successful</h1>
      <p><strong>Refresh Token:</strong> ${refresh_token}</p>
      <p><strong>Access Token:</strong> ${access_token}</p>
    `;

    return { statusCode: 200, body: successPage };
  } catch (error) {
    return {
      statusCode: 500,
      body: \`Token exchange failed: \${error.response?.data?.error_description || error.message}\`,
    };
  }
};