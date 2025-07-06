const axios = require('axios');

exports.handler = async (event) => {
  const refreshToken = event.queryStringParameters.refresh_token;

  if (!refreshToken) {
    return {
      statusCode: 400,
      body: 'Refresh token missing'
    };
  }

  try {
    const tokenResponse = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', new URLSearchParams({
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,  // 🔑 Using client secret
      scope: 'offline_access https://graph.microsoft.com/Mail.Read',
      grant_type: 'refresh_token',
      refresh_token: refreshToken
    }).toString(), {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });

    const { access_token, refresh_token: newRefreshToken } = tokenResponse.data;

    return {
      statusCode: 200,
      body: JSON.stringify({ access_token, refresh_token: newRefreshToken })
    };

  } catch (error) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: error.response ? error.response.data : error.message })
    };
  }
};
