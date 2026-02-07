# etc-copilot (Chat – SharePoint Web Part)

**Repo:** [github.com/pablomferrari/etc-copilot](https://github.com/pablomferrari/etc-copilot)

A ChatGPT-style chat web part for SharePoint that uses OpenAI. The API key can be stored in the web part properties (for a selected group of managers) or later in a config list.

## Prerequisites

- Node.js 18.x or 20.x (build uses `NODE_OPTIONS=--openssl-legacy-provider` for Node 18)
- SharePoint Online (or workbench for local testing)
- OpenAI API key

**If `npm install` or build fails:** Some antivirus tools rename files in `node_modules` (e.g. `db.json` → `db.json.DELETE.*`). Exclude `C:\dev\etc\chat-2etc\node_modules` from real-time scanning, or restore the renamed files and run again.

## Setup

1. **Install dependencies**

   ```bash
   npm install
   ```

2. **Build**

   ```bash
   npm run build
   ```

3. **Run locally (workbench)**

   Update `config/serve.json` with your SharePoint site URL, then:

   ```bash
   npm run serve
   ```

   Open the hosted workbench URL and add the **Chat** web part to the page.

4. **Set the API key**

   - Edit the page, select the Chat web part, click the pencil (Edit web part).
   - In the property pane, enter your **OpenAI API key** and confirm.

## Deploy to SharePoint

1. **Package the solution**

   ```bash
   npm run package
   ```

2. **Upload** the `.sppkg` from `sharepoint/solution/` to your tenant app catalog.

3. **Add the app** to a site, then add the **Chat** web part to a page and set the API key in the property pane.

## Project structure

- `src/webparts/chat/` – Web part entry and manifest
- `src/webparts/chat/components/` – React chat UI (messages, input, send)
- `config/` – SPFx build and deploy config

## Plan

See [PLAN.md](./PLAN.md) for the roadmap (routing, Ollama, Azure Function).
