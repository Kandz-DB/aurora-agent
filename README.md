Aurora — R2S Project Management Intelligence
Files
server.js     — everything: API, AI, Monday sync, batch, cost controls
public/index.html — the full Aurora portal UI
prompt.txt    — Aurora's identity and instructions
package.json  — dependencies
render.yaml   — Render deployment config
Deploy to Render
Push this repo to GitHub
Connect to Render as a Web Service
Build: npm install | Start: node server.js
Add environment variables (see below)
Environment variables
ANTHROPIC_API_KEY    — from console.anthropic.com
MONDAY_API_KEY       — from Monday.com profile > Developers
MONDAY_BOARD_ID      — number in your Monday.com board URL
MONTHLY_SPEND_CAP_USD — default 20 (USD)
FRONTEND_URL         — your Render URL (e.g. https://aurora-r2s.onrender.com)
