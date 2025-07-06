const axios = require('axios');

exports.handler = async function (event, context) {
  const params = event.queryStringParameters || {};
  const code = params.code;
  const refreshToken = params.refresh_token;

  try {
    let tokenResponse;

    if (code) {
      // Exchange authorization code for tokens WITHOUT client_secret
      const formData = new URLSearchParams();
      formData.append('client_id', process.env.CLIENT_ID);
      formData.append('scope', 'offline_access https://graph.microsoft.com/Mail.Read');
      formData.append('code', code);
      formData.append('redirect_uri', process.env.REDIRECT_URI);
      formData.append('grant_type', 'authorization_code');

      tokenResponse = await axios.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        formData.toString(),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
      );

    } else if (refreshToken) {
      // Refresh token request WITHOUT client_secret
      const formData = new URLSearchParams();
      formData.append('client_id', process.env.CLIENT_ID);
      formData.append('scope', 'offline_access https://graph.microsoft.com/Mail.Read');
      formData.append('refresh_token', refreshToken);
      formData.append('grant_type', 'refresh_token');

      tokenResponse = await axios.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        formData.toString(),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
      );

    } else {
      return {
        statusCode: 400,
        body: 'Missing code or refresh_token parameter',
      };
    }

    const { access_token, refresh_token } = tokenResponse.data;

    return {
      statusCode: 200,
      body: JSON.stringify({
        access_token,
        refresh_token,
      }),
    };

  } catch (error) {
    return {
      statusCode: 500,
      body: `Token exchange failed: ${error.response?.data?.error_description || error.message}`,
    };
  }
};
