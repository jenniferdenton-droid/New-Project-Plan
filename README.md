# Moxie Ops Project Plans — Deploy Bundle

These 4 files are everything you need to deploy the AI Assistant-powered project tracker to Vercel.

## Contents

```
deploy-bundle/
├── index.html              ← Main app (HTML + JS + CSS, single file)
├── vercel.json             ← Vercel config (60s function timeout, 1GB memory)
├── package.json            ← Node dependencies (Anthropic SDK, googleapis)
├── .gitignore              ← Prevents node_modules from being committed
└── api/
    └── ai-assistant.js     ← Serverless function (Claude + Slack + Gmail)
```

## How to Deploy

### Prerequisite: Environment Variables (Already Done)

These 6 should be set in Vercel → Settings → Environment Variables:
- `ANTHROPIC_API_KEY`
- `SLACK_BOT_TOKEN`
- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `GOOGLE_REFRESH_TOKEN`
- `GMAIL_FROM_EMAIL`

### Push to GitHub (auto-deploys to Vercel)

```bash
# From your repo root
cp ~/path/to/deploy-bundle/index.html .
cp ~/path/to/deploy-bundle/vercel.json .
cp ~/path/to/deploy-bundle/package.json .
cp ~/path/to/deploy-bundle/.gitignore .
mkdir -p api && cp ~/path/to/deploy-bundle/api/ai-assistant.js api/

git add index.html vercel.json package.json .gitignore api/ai-assistant.js
git commit -m "Add AI Assistant + Workstreams + Gmail/Slack integration"
git push origin main
```

Or use the GitHub web UI to upload each file in the right location.

## What's New in This Build

- 🤖 AI Assistant tab — process transcripts, send status updates, quick Slack/email
- 🧭 Workstreams tab — tasks grouped by team (Sales, PSM, Onboarding, etc.)
- 👥 Dashboard "Owners by Team" card
- 🧩 Process Flow tab with Mermaid diagrams
- 📄 Existing Google Doc link in Proposal Builder
- Inline phase dropdown on tasks/AIs (move items between workstreams)
- Editable due dates and priorities on action items in Plan tab
- Fixed: duplicate "BookKeeping" project on refresh bug

## Smoke Test After Deploy

1. Open your Vercel URL → click 🤖 AI Assistant tab
2. Fill in Settings: Slack channel, email recipients
3. Test in order:
   - Quick Slack message
   - Quick Email
   - Process Transcript (paste a 3-line fake transcript)
   - Send Status Update

If any fails, check Vercel → Functions → ai-assistant → live logs for error details.

## Cost

- Vercel: Hobby plan (free) is fine — Generate Deck is hidden because it needs 60s timeout (Pro plan)
- Anthropic: ~$0.02 per transcript/status call · free $5 starter credit covers ~50 calls
- Slack + Gmail: free
- Estimated: ~$8/mo total once you're using it daily
