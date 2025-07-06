# Microsoft OAuth on Netlify

Simple static site with Netlify Functions to authenticate with Microsoft OAuth and retrieve refresh tokens.

## Deployment

1. Push this project to GitHub.
2. Connect the repository to Netlify.
3. Set environment variables:
   - `CLIENT_ID`
   - `CLIENT_SECRET`
   - `REDIRECT_URI` (example: `https://your-netlify-site.netlify.app/.netlify/functions/token`)
4. Done! The OAuth flow is fully functional.