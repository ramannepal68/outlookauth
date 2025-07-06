exports.handler = async function (event, context) {
  const clientId = process.env.CLIENT_ID;
  const redirectUri = process.env.REDIRECT_URI;
  const scope = encodeURIComponent('offline_access https://graph.microsoft.com/Mail.Read');

  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${redirectUri}&response_mode=query&scope=${scope}`;

  return {
    statusCode: 302,
    headers: {
      Location: authUrl,
    },
  };
};