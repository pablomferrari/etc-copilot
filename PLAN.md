# Plan: SharePoint ChatGPT-like App with Smart Model Routing

## Goal

Build a SharePoint app that behaves like ChatGPT while **minimizing API costs** by:
1. **Classifying** each user request (simple vs. complex)
2. **Routing** to the cheapest suitable model first (e.g. small OpenAI models or later Ollama)
3. **Reserving** expensive models (e.g. GPT-4) only for complex tasks
4. **Future**: Route simple tasks to on-prem **Ollama**, complex to **OpenAI**

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────────┐
│  SharePoint (SPFx Web Part or Teams Tab)                                 │
│  ┌───────────────────────────────────────────────────────────────────┐  │
│  │  Chat UI (messages, input, history)                                │  │
│  └───────────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────┐
│  Your Backend (Azure Function / Node API / .NET API)                     │
│  • Keeps OpenAI key server-side (never in SharePoint)                    │
│  • Request classifier (simple vs complex)                                 │
│  • Router → Ollama (future) or OpenAI                                    │
└─────────────────────────────────────────────────────────────────────────┘
        │                                    │
        ▼                                    ▼
┌───────────────────┐              ┌───────────────────┐
│  Ollama (on-prem)  │              │  OpenAI API        │
│  • Simple Q&A      │              │  • gpt-4o-mini     │
│  • Summarization   │              │  • gpt-4o (complex)│
└───────────────────┘              └───────────────────┘
```

---

## Simplest Start: Key in SharePoint (No Backend Yet)

For a **limited audience** (e.g. a selected group of managers), you can keep things minimal and store the OpenAI key in SharePoint. No backend until you add Ollama and routing.

**Architecture now (no backend):**

```
┌─────────────────────────────────────────────────────────────────────────┐
│  SharePoint (SPFx Web Part)                                              │
│  • Chat UI (messages, input, history)                                    │
│  • OpenAI key from: SharePoint list, or web part properties, or .env   │
│  • Calls OpenAI API directly from the browser                            │
└─────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌───────────────────┐
│  OpenAI API        │
│  (e.g. gpt-4o-mini)│
└───────────────────┘
```

**Where to put the key in SharePoint:**

- **Option A – SharePoint list**: Create a config list (e.g. "Chat Config") with restricted permissions (managers only). One item with a column for the API key. SPFx reads it at load time via the list API. Key is in SharePoint data, not in the deployed package.
- **Option B – Web part property**: Add an "API key" property to the web part. Admins enter it when adding the web part to a page. Stored in the page; only people who can edit the page see it.
- **Option C – Build-time env**: Put the key in a `.env` or config file used at build time (e.g. in a private repo). It gets baked into the bundle — acceptable only if the app is deployed to a closed audience and you accept that the key is in the JS.

**Tradeoff**: The key is visible to anyone who can inspect the page/source/network (e.g. DevTools). For a small, trusted group of managers and with OpenAI usage limits set, this is a reasonable choice. When you add Ollama and an Azure Function later, you move the key to the backend and the app stops using a client-side key.

**Later**: When your Ollama server is ready, introduce an Azure Function (or other backend), move the key there, add classifier + router, and change the SPFx app to call your `/api/chat` instead of OpenAI. The UI stays the same.

### When to move file handling to the backend (Azure Function)

Right now, **document text extraction (PDF, Word, Excel) runs in the browser**. That keeps the current “no backend” setup but has tradeoffs:

| Approach | Pros | Cons |
|----------|------|------|
| **Client-side (current)** | No backend; key in SharePoint; works for Word/Excel/text in many environments. | PDF extraction can fail in some SharePoint/browser contexts (CSP, iframe, PDF.js quirks). User may see “extraction failed” and the model saying it can’t access attachments. |
| **Backend (Azure Function)** | Reliable PDF (and other doc) extraction (e.g. Node with `pdf-parse` or PDF.js); API key and routing live server-side; same endpoint for chat + files. | Requires hosting, CORS, and auth (e.g. SharePoint token or API key). |

**Recommendation:**

1. **Try the current architecture first**  
   After the latest fixes you get:
   - **Visible extraction status**: attachment bar shows “text ready” or “extraction failed” per file.
   - **Clear model behavior**: if extraction failed, the system prompt tells the model not to say “I can’t access attachments” but to suggest pasting text or using a backend.

2. **If PDFs (or other docs) still often fail** in your SharePoint/workbench environment, **move to an Azure Function**:
   - SPFx sends **files (or base64)** to your function instead of extracting in the browser.
   - The function extracts text (e.g. Node `pdf-parse` or PDF.js), builds the user message, calls OpenAI (and later your router), returns the reply.
   - API key and any future routing (Ollama vs OpenAI) stay in the function; the web part only calls `POST /api/chat` (with optional multipart or JSON body for files).

So: **fix with current architecture first** (you now have better feedback and model instructions). If PDF extraction remains unreliable in production, **then** move file handling to a backend (e.g. Azure Function) for a more reliable solution.

---

## Alternative: Backend from Day One

**Do we need a backend now?**  
You need *something* that holds the OpenAI key off the client. Calling OpenAI directly from the SharePoint app would expose the key in the browser (network tab, bundled JS). So a thin backend is required from day one — but it can be very small.

**Sensible approach:**

| When | What you have | Backend’s job |
|------|----------------|---------------|
| **Now** | SharePoint chat UI + one place that holds the key | Minimal: one endpoint that proxies requests to OpenAI (e.g. always `gpt-4o-mini`). No classifier, no router. |
| **Later** (Ollama ready) | Same UI + AI server with Ollama | Move/evolve that proxy into an **Azure Function**: add classifier + router; simple → Ollama, complex → OpenAI. |

So: **start with a minimal backend** (proxy only, single model). When your Ollama server is in place, **move the logic into an Azure Function** and add routing there. The SharePoint app keeps calling “one chat API”; only the backend implementation changes.

- **Now**: Minimal backend (any host: Azure Function, small Node/Express app, or even a serverless proxy). Single model, no routing.
- **Later**: Implement or migrate to Azure Function and add classifier + router + Ollama + fallback.

The rest of this plan (classification, routing, Ollama) applies when you introduce or evolve that backend.

---

## Phase 1: Foundation (SharePoint + Backend + OpenAI Only)

### 1.1 Backend API (required first)

- **Role (minimal now)**: Proxy for OpenAI and hold API key. No classifier or router yet.
- **Role (later, when Ollama is ready)**: Same proxy, but add classifier + router in an Azure Function.
- **Options**:
  - **Azure Functions** (Node or C#) – good when you add routing; you can start with a single HTTP-triggered function.
  - **Node/Express or .NET API** – fine for “minimal now”; migrate to Azure Function later if you prefer.
- **Endpoints**:
  - `POST /api/chat` – receives messages from SharePoint, returns assistant reply (streamed or not).
  - *(Later)* Optional: `POST /api/classify` for testing.
- **Responsibilities now**: Validate request (e.g. auth token), call OpenAI (e.g. one model like `gpt-4o-mini`), return response.
- **Responsibilities later**: Add classification (see §2) and routing to Ollama / OpenAI (see §3).

### 1.2 SharePoint app (SPFx)

- **Type**: SPFx 1.18+ Web Part (or Teams tab using same SPFx app).
- **UI**:
  - Chat message list (user + assistant).
  - Input box + send button.
  - Optional: “force model” dropdown for power users (e.g. “Always use GPT-4”).
- **Calls**: Only your backend `POST /api/chat` (no OpenAI key in SPFx).
- **Auth**: Use SharePoint/Teams context; send token (e.g. Bearer) to backend so only your org can call the API.

### 1.3 Security

- Store OpenAI key in **Azure Key Vault** or backend **env vars** (never in repo or client).
- Backend checks **Azure AD / SharePoint token** and only accepts requests from your tenant (and optionally specific sites).

---

## Phase 2: Request Classification (Minimize Cost)

Classify **before** calling any model. Use a very cheap call to decide the route.

### Option A – Rule-based (no extra API call)

- **Heuristics** (e.g. in backend):
  - Short message + question mark → likely simple (e.g. “What is X?”).
  - Long message, bullet points, “analyze”, “compare”, “code”, “debug” → likely complex.
  - Keywords: “summarize”, “translate one sentence” → simple; “explain step by step”, “write a function” → complex.
- **Pros**: No cost, fast. **Cons**: Less accurate.

### Option B – Classifier model (one cheap call)

- Use **one** small, cheap call to decide: “simple” vs “complex”.
- **Model**: e.g. **gpt-4o-mini** with a tiny system prompt:
  - “Classify the user message: reply with only SIMPLE or COMPLEX. SIMPLE = factual Q&A, short summary, simple rewrite. COMPLEX = coding, long analysis, multi-step reasoning, creative writing.”
- **Pros**: More accurate routing, still cheap (one mini call per turn). **Cons**: Slight latency and small cost.

**Recommendation**: Start with **Option A**; add **Option B** if you see mis-routes (e.g. too many complex queries on cheap model or vice versa).

---

## Phase 3: Model Routing (OpenAI Only, Then + Ollama)

### 3.1 OpenAI-only routing (now)

- **Simple**:
  - **gpt-4o-mini** (or **gpt-3.5-turbo** if you want even cheaper).
- **Complex**:
  - **gpt-4o** (or **gpt-4-turbo**), or **gpt-4o-mini** first and “escalate” to gpt-4o only if needed (e.g. retry with stronger model on certain failures).
- Store mapping in config (env or Key Vault):
  - `SIMPLE_MODEL=gpt-4o-mini`
  - `COMPLEX_MODEL=gpt-4o`

### 3.2 Add Ollama (later)

- **Ollama** runs on-prem (or in your network). No per-token cost.
- **Router change**:
  - **Simple** → call your **Ollama** API (e.g. `llama3.2`, `phi3`, `mistral`) instead of OpenAI.
  - **Complex** → still **OpenAI** (gpt-4o-mini or gpt-4o).
- Backend must:
  - Have network access to Ollama (e.g. `http://your-ollama-server:11434/api/chat`).
  - Map Ollama’s request/response format to your internal “chat” format so the SharePoint UI stays unchanged.
- **Fallback**: If Ollama is down or times out, route simple requests to **gpt-4o-mini** so the app still works.

---

## Phase 4: Implementation Checklist

### Start simple (now)

| # | Task | Notes |
|---|------|--------|
| 1 | Create minimal backend (Azure Function or small Node/Express app) | Single project; only needs to proxy chat |
| 2 | Add `POST /api/chat`: accept messages, call OpenAI (e.g. `gpt-4o-mini`), return reply | Use OpenAI SDK; key from env or Key Vault |
| 3 | Add auth: validate SharePoint/AD token (optional but recommended) | So only your tenant can call the API |
| 4 | Scaffold SPFx 1.18+ web part | Yeoman generator |
| 5 | Build chat UI (message list + input), call backend `/api/chat` | No OpenAI key in client |
| 6 | Deploy backend and SPFx; test end-to-end | |

### When Ollama is ready (later)

| # | Task | Notes |
|---|------|--------|
| 7 | Move/implement backend as Azure Function (if not already) | Same `/api/chat` contract |
| 8 | Implement classifier (rules or one cheap LLM call) | Output: SIMPLE \| COMPLEX |
| 9 | Implement router: SIMPLE → Ollama, COMPLEX → OpenAI | Config-driven; fallback simple → gpt-4o-mini if Ollama down |
| 10 | Add Ollama client: call `http://your-ollama:11434/api/chat`, map request/response | UI unchanged |
| 11 | (Optional) “Force model” in UI for testing | |

---

## Tech Stack Summary

| Layer | Suggested stack |
|-------|------------------|
| SharePoint app | SPFx 1.18+, React, Fluent UI (optional) |
| Backend | Node (Azure Functions or Express) or C# (Azure Functions or ASP.NET Core) |
| Auth | Azure AD / SharePoint context; validate JWT in backend |
| Secrets | Azure Key Vault or env vars |
| OpenAI | Official SDK (openai npm or OpenAI .NET); key server-side only |
| Ollama (later) | HTTP client to `http://ollama-host:11434/api/chat` |

---

## Cost Control Tips

1. **Set usage caps** in OpenAI dashboard (e.g. monthly limit).
2. **Log** which model served each request (simple/complex, model name) for analytics.
3. **Tune classifier** so “complex” is only when truly needed (avoid over-use of gpt-4o).
4. **Cache** common answers (e.g. company FAQs) in backend or SharePoint list to skip LLM calls when possible.
5. **Stream** responses so users see output early; use same token limits for both routes to avoid runaway length.

---

## Next Steps

**Now (start simple):**
1. Choose a minimal backend host (e.g. one Azure Function, or a small Node/Express app).
2. Implement a single `POST /api/chat` that proxies to OpenAI (one model, e.g. `gpt-4o-mini`).
3. Scaffold the SPFx chat web part and point it at your `/api/chat` URL.
4. Deploy and test. No classifier or router yet.

**Later (when you have Ollama):**
5. Move or reimplement that logic in an Azure Function.
6. Add classifier + router; wire simple → Ollama, complex → OpenAI, with fallback.

If you tell me your preferred stack (Node vs C#, Azure Function vs Express now), I can outline concrete file structure and code for the minimal backend and SPFx next.
