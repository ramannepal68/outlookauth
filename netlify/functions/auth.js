const CLIENT_ID = process.env.CLIENT_ID;
const REDIRECT_URI = process.env.REDIRECT_URI;
const SCOPES = 'offline_access https://graph.microsoft.com/Mail.Read';

exports.handler = async (event) => {
  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${CLIENT_ID}&scope=${encodeURIComponent(SCOPES)}&response_type=code&redirect_uri=${encodeURIComponent(REDIRECT_URI)}&response_mode=query`;

  return {
    statusCode: 302,
    headers: {
      Location: authUrl
    },
    body: ''
  };
};
