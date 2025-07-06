const axios = require('axios');

exports.handler = async (event) => {
  const code = event.queryStringParameters.code;

  if (!code) {
    return {
      statusCode: 400,
      body: 'Authorization code missing'
    };
  }

  try {
    const tokenResponse = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', new URLSearchParams({
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,  // 🔑 Using client secret
      scope: 'offline_access https://graph.microsoft.com/Mail.Read',
      grant_type: 'authorization_code',
      redirect_uri: process.env.REDIRECT_URI,
      code: code
    }).toString(), {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });

    const { access_token, refresh_token } = tokenResponse.data;

    return {
      statusCode: 200,
      body: JSON.stringify({ access_token, refresh_token })
    };

  } catch (error) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: error.response ? error.response.data : error.message })
    };
  }
};
