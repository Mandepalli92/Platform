# My Microsoft 365 Assistant
Sign in with Microsoft, ask a question, and get answers from your own mail, files and calendar with citations.

## Setup
1. **Entra app**: in Azure Portal > Entra ID > App registrations > New. Use platform: Single-page application. Add the redirect URI for your GitHub Pages site, for example `https://<your-user>.github.io/<repo>/`. Grant delegated permissions: User.Read, Mail.Read, Files.Read.All, Calendars.Read. Copy the Application (client) ID into `config.js`.
2. **Worker**: edit `wrangler.toml` and set `ORIGIN` to your Pages URL. Leave `ALLOWED_USER` blank to allow the signed-in Microsoft account, or set it to one email to lock the app to a single account. Deploy with `npx wrangler deploy` and store the LLM key with `npx wrangler secret put ANTHROPIC_API_KEY`. Put the resulting worker URL in `config.js`.
3. **Pages**: push to GitHub and publish the site in GitHub Pages or use the included workflow.

## Notes
- If the app is in a work tenant, admin consent may be required. Mail and file content is sent to the model provider, so confirm your company policy before using it.
- The app reads from the signed-in user’s Microsoft 365 data, not a shared account, when `ALLOWED_USER` is left empty.
- The worker has no rate limit. Add one before exposing it publicly.
