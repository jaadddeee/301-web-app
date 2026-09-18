# Redirect Map Builder — hosted version

Browser-based front end for `sitemap_to_redirect_map.py` and `match_new_site.py`.
It doesn't reimplement the scraping — it runs those two scripts as background
subprocesses (same as the old desktop GUI did) and streams their output into
the browser.

## Run locally

```
pip install -r requirements.txt
python app.py
```
Open http://localhost:5000

## Deploy (hosted, for teammates/clients)

Use a host that runs a **persistent process** — Render, Railway, Fly.io, or a
small VPS all work well. Avoid serverless platforms (Vercel/Netlify functions):
crawls can take minutes on slow sites, which exceeds their request timeouts,
and they don't give you a writable persistent process for background jobs.

1. Push this folder to a git repo.
2. On your host, set the start command to what's in `Procfile`:
   `gunicorn -w 1 --threads 8 -k gthread app:app`
   (`-w 1` is required — job state lives in memory in this one process.)
3. Set environment variables:
   - `APP_PASSWORD` — a shared password. **Set this before making it public** —
     without it, anyone with the link can use the tool to crawl sites through
     your server.
   - `SECRET_KEY` — any random string, so login sessions survive a restart.
4. Deploy. Share the URL + password with your team/clients.

## Notes

- Each build/match job gets its own folder under `jobs/`, auto-deleted after
  6 hours.
- The app rejects URLs that resolve to a private/internal IP address (basic
  protection against the tool being pointed at your own server's internal
  network).
- If you outgrow one worker process (e.g. want to scale to multiple
  dynos/instances), job state will need to move from the in-memory `jobs`
  dict to something shared like Redis — flag it if you get there and it's a
  small change.
# 301-web-app
