const axios = require('axios');

exports.handler = async function (event, context) {
  const params = event.queryStringParameters;
  const code = params.code;

  if (!code) {
    return { statusCode: 400, body: 'No code provided' };
  }

  try {
    // Build the form data correctly for application/x-www-form-urlencoded
    const formData = new URLSearchParams();
    formData.append('client_id', process.env.CLIENT_ID);
    formData.append('scope', 'offline_access https://graph.microsoft.com/Mail.Read');
    formData.append('code', code);
    formData.append('redirect_uri', process.env.REDIRECT_URI);
    formData.append('grant_type', 'authorization_code');
    formData.append('client_secret', process.env.CLIENT_SECRET);

    // Send the token request with form-encoded body
    const tokenResponse = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', formData.toString(), {
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
      body: `Token exchange failed: ${error.response?.data?.error_description || error.message}`,
    };
  }
};
