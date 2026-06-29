<div align="center">

# 📄 APA Paper Generator

**A FastAPI service that generates fully formatted APA 7th edition academic papers via AI and exports them as `.docx` files.**

[![Python](https://img.shields.io/badge/Python-3.12-3776AB?logo=python&logoColor=white)](https://python.org)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.138-009688?logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com)
[![Docker](https://img.shields.io/badge/Docker-ready-2496ED?logo=docker&logoColor=white)](./Dockerfile)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)

Fill in your assignment details, press **Generate**, and receive a complete, properly structured APA 7 paper — title page, abstract, body sections with in-text citations, and a references page — ready to download as `.docx`.

</div>

---

## ✨ Features

- **APA 7th edition compliance** — title page, abstract, Level 1 headings, hanging-indent references, 1-inch margins, Times New Roman 12pt double-spaced
- **AI-powered generation** — sends your prompt to VoidAI (gpt-4o by default); supports any OpenAI-compatible endpoint including LiteLLM proxies
- **`.docx` export** — built server-side with `python-docx`; no third-party Word services
- **Live browser preview** — a single-file `static/index.html` renders the paper in an APA-styled preview before you download
- **Agent-friendly endpoint** — `/api/agent/generate` returns JSON + Markdown + a download URL in one call
- **Configurable** — word count (500–2000), source count (2–5), tone (Academic / Reflective / Analytical), and model are all runtime config
- **Healthcheck endpoint** — `/api/health` for Docker and reverse-proxy liveness probes

---

## 🧱 Tech Stack

| Layer            | Technology                                    |
| ---------------- | --------------------------------------------- |
| Language         | Python 3.12                                   |
| Web framework    | FastAPI                                       |
| ASGI server      | Uvicorn                                       |
| HTTP client      | httpx (async)                                 |
| Document export  | python-docx                                   |
| Data validation  | Pydantic v2                                   |
| AI provider      | VoidAI → gpt-4o (OpenAI-compatible API)       |
| Frontend         | Vanilla HTML/CSS/JS (single static file)      |
| Container        | Docker + Docker Compose + Watchtower          |

---

## 🚀 Quick Start

### Local development

**Prerequisites:** Python 3.12, pip

```bash
# 1. Clone
git clone https://github.com/The-Code-Labz/apa-paper-generator.git
cd apa-paper-generator

# 2. Create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Configure environment
cp .env.example .env
# → edit .env and set SHARED_VOID_API_FREE_KEY

# 5. Run the dev server
uvicorn app:app --reload --port 8000
```

Open [http://localhost:8000](http://localhost:8000) in your browser.

---

### Docker (recommended for production)

```bash
# 1. Configure environment
cp .env.example .env
# → edit .env and set SHARED_VOID_API_FREE_KEY

# 2. Build and start
docker compose up -d --build

# 3. Check it's healthy
docker compose ps
curl http://localhost:8030/api/health
```

The service is now available at **http://localhost:8030**.

Generated `.docx` files are stored in the named Docker volume `apa-papers` and persist across container restarts.

---

## 🔑 Environment Variables

Create a `.env` file at the project root (never commit it — it is in `.gitignore`).

| Variable                  | Default                          | Required | Description                                                       |
| ------------------------- | -------------------------------- | -------- | ----------------------------------------------------------------- |
| `SHARED_VOID_API_FREE_KEY` | —                               | ✅       | VoidAI API key. Get one at [voidai.app](https://voidai.app)       |
| `APA_MODEL`               | `gpt-4o`                         | No       | Model name forwarded to the AI provider                           |
| `VOID_API_BASE`           | `https://api.voidai.app/v1`      | No       | API base URL. Point at a LiteLLM proxy for internal routing       |

**Example `.env`:**

```env
SHARED_VOID_API_FREE_KEY=sk-your-key-here
APA_MODEL=gpt-4o
# VOID_API_BASE=http://your-litellm-host:4000
```

---

## 📡 API Reference

All request bodies are JSON (`Content-Type: application/json`).

### GET `/api/health`

Liveness probe. Returns `200 OK` when the service is running.

```json
{ "ok": true, "model": "gpt-4o", "key": true }
```

---

### POST `/api/generate`

Generate a paper and return the structured JSON. Use this when you want to render the paper yourself or pass it to `/api/docx`.

**Request body:**

```json
{
  "title":       "The Impact of Sleep on Academic Performance",
  "author":      "Jane Doe",
  "course":      "PSY-102: General Psychology",
  "instructor":  "Dr. Smith",
  "institution": "Grand Canyon University",
  "due_date":    "June 29, 2026",
  "prompt":      "Discuss how sleep deprivation affects student GPA and cognitive performance.",
  "word_count":  1000,
  "sources":     3,
  "tone":        "Academic"
}
```

| Field         | Type    | Default                    | Notes                              |
| ------------- | ------- | -------------------------- | ---------------------------------- |
| `title`       | string  | —                          | Required                           |
| `author`      | string  | —                          | Required                           |
| `course`      | string  | —                          | Required                           |
| `instructor`  | string  | —                          | Required                           |
| `institution` | string  | `Grand Canyon University`  | Optional                           |
| `due_date`    | string  | —                          | Required                           |
| `prompt`      | string  | —                          | Required — the assignment topic    |
| `word_count`  | integer | `1000`                     | Target body word count             |
| `sources`     | integer | `3`                        | Number of references to generate   |
| `tone`        | string  | `Academic`                 | `Academic` / `Reflective` / `Analytical` |

**Response:** A JSON object with `title_page`, `abstract`, `sections[]`, and `references[]`.

---

### POST `/api/docx`

Convert a paper JSON object (from `/api/generate`) into a properly formatted `.docx` download.

**Request body:** the JSON returned by `/api/generate`.

**Response:** `application/vnd.openxmlformats-officedocument.wordprocessingml.document` binary stream with `Content-Disposition: attachment`.

---

### POST `/api/agent/generate`

One-shot endpoint for agent/automation use. Generates the paper, builds the `.docx`, saves it to disk, and returns everything in a single response.

**Request body:** same as `/api/generate`.

**Response:**

```json
{
  "paper":        { /* same structure as /api/generate */ },
  "markdown":     "# Title\n\n...",
  "download_url": "/api/download/550e8400-e29b-41d4-a716-446655440000",
  "paper_id":     "550e8400-e29b-41d4-a716-446655440000"
}
```

---

### GET `/api/download/{paper_id}`

Download a previously generated `.docx` file by UUID.  
Returns `404` if the file does not exist or if the path escapes the allowed directory (path traversal protection is enforced).

---

### GET `/`

Serves `static/index.html` — the browser-based form and live paper preview.

---

## 🐳 Docker Details

### Ports

| Host   | Container | Protocol |
| ------ | --------- | -------- |
| `8030` | `8000`    | HTTP     |

### Volumes

| Volume       | Mount path        | Purpose                              |
| ------------ | ----------------- | ------------------------------------ |
| `apa-papers` | `/tmp/apa-papers` | Persists generated `.docx` files     |

### Auto-updates (Watchtower)

The `docker-compose.yml` includes a [Watchtower](https://containrrr.dev/watchtower/) service that polls the container registry every **5 minutes** and automatically restarts any container whose image has been updated.

To enable auto-update for the `apa-generator` container, add this label to it in `docker-compose.yml`:

```yaml
labels:
  com.centurylinklabs.watchtower.enable: "true"
```

To disable Watchtower entirely, comment out the `watchtower` service block.

### Rebuilding after code changes

```bash
docker compose up -d --build
```

The Dockerfile is structured so that dependency installation (`pip install`) is cached in a separate layer — code-only changes rebuild only the final layer and are very fast.

---

## 📁 Project Structure

```
apa-paper-generator/
├── app.py               # FastAPI application — routes, AI call, docx builder
├── requirements.txt     # Python dependencies (pinned)
├── Dockerfile           # Multi-layer build; non-root user; HEALTHCHECK
├── docker-compose.yml   # Service + Watchtower + named volume + network
├── .env.example         # Environment variable template
├── .gitignore
├── .dockerignore
└── static/
    └── index.html       # Single-file frontend (vanilla JS, no build step)
```

---

## 🔒 Security Notes

### Container runs as a non-root user

The `Dockerfile` creates a dedicated `apa` system user. The process runs as this user inside the container, limiting blast radius if the application is ever exploited.

### Never commit `.env`

`.env` is listed in both `.gitignore` and `.dockerignore`. The only file committed to the repository is `.env.example`, which contains no real credentials.

### No rate limiting

The `/api/generate` endpoint makes an upstream AI call that takes up to 60 seconds. There is currently **no rate limiting** on this endpoint. If the service is exposed publicly, add rate limiting at the reverse-proxy layer (nginx `limit_req`, Traefik middleware, etc.) before going to production.

### Path traversal protection

`/api/download/{paper_id}` validates that the resolved file path stays within the allowed directory before serving the file.

---

## 🤝 Contributing

1. Fork the repository.
2. Create a feature branch: `git checkout -b feat/your-feature`
3. Install dependencies: `pip install -r requirements.txt`
4. Make your changes and test locally with `uvicorn app:app --reload`.
5. Open a Pull Request against `main` with a clear description.

### Code style

- Python type hints everywhere.
- Keep AI logic in `generate_paper()`, docx logic in `build_docx()`.
- No new external services without updating `.env.example` and this README.

---

## 📄 License

This project is licensed under the **MIT License**.

---

<div align="center">

Built with care by [The-Code-Labz](https://github.com/The-Code-Labz)

</div>
