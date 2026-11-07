# My Microsoft 365 Assistant
Sign in with Microsoft, ask a question, get an answer drawn only from your own mail, files and calendar, with links.

## Setup
1. **Entra app**: portal.azure.com > Entra ID > App registrations > New. Platform: *Single-page application*, redirect URI `https://<user>.github.io/<repo>/`. API permissions (delegated): User.Read, Mail.Read, Files.Read.All, Calendars.Read. Copy the Application (client) ID into `config.js`.
2. **Worker**: edit `worker/wrangler.toml` (ALLOWED_USER = your sign-in email, ORIGIN = `https://<user>.github.io`). Then `cd worker && npx wrangler deploy && npx wrangler secret put ANTHROPIC_API_KEY`. Put the worker URL in `config.js`.
3. **Pages**: push to GitHub (`main`), then Settings > Pages > Source: GitHub Actions.

## Caveats
- Work tenants often block third-party apps or require admin consent. Mail and file content is sent to the LLM provider: check your employer's policy first.
- Untested against a live tenant. Graph search behaviour can differ between work and personal accounts.
- The worker has no rate limit. Add one before sharing the URL.
