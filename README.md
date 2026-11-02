# Pekora — Implementation Plan

> **Product:** Personal, local-first AI presentation generator.  
> **Canonical state:** Pekora slide JSON.  
> **Canva role:** Optional final handoff for further visual editing; Canva changes do not sync back into Pekora.  
> **Code status:** Representative implementation code and contracts—not a finished application.

---

## 1. Product definition

Pekora turns a topic, pasted text, URL, PDF, or images into a presentation. It offers quick and guided creation, fixed layouts and themes, limited corrections, image approval before use, speaker scripts, and exports to PPTX, PDF, and PNG.

The app supports multiple cloud-model providers and a local Ollama option when Ollama is already running. It never auto-installs models and it does not generate images.

### Core rules

1. **Deck JSON is the source of truth.** UI changes modify validated JSON; PPTX/PDF/PNG are rebuilt artifacts.
2. **AI writes content, not layouts.** Models select only from Pekora-owned layout IDs and schema fields.
3. **Images require approval.** Search results are never included until the user explicitly approves them.
4. **Canva is a one-way handoff.** Pekora exports/imports a finished PPTX to Canva, then opens Canva for human editing. Canva changes are a separate Canva version.
5. **Offline failure is graceful.** Without network or Ollama, Pekora can still create a basic deck from headings and paragraphs.

---

## 2. Architecture

```text
React PWA
  ├─ input / brief / theme settings
  ├─ outline and image approval gallery
  ├─ small corrective editor
  ├─ export controls
  └─ Canva Connect + Open in Canva
              │
              ▼
FastAPI
  ├─ ingest: text, PDF, OCR, URL
  ├─ AI router: cloud providers + OpenRouter + optional Ollama
  ├─ schema validation and fallback policy
  ├─ fixed-layout PPTX builder
  ├─ LibreOffice PDF + PyMuPDF PNG renderer
  ├─ approved-image service
  ├─ speaker-script service
  └─ Canva OAuth + import handoff service
              │
       ┌──────┼────────┬──────────────┐
       ▼      ▼        ▼              ▼
    Models  Ollama   Image sources   Canva REST APIs
                                      ├─ OAuth PKCE
                                      ├─ PPTX design import
                                      └─ optional design generation (preview)
```

### Canva integration boundary

**Recommended v1:** build in Pekora, export the completed PPTX, import it into the user’s Canva account, and show the resulting Canva edit link. Canva’s import API supports PPTX and returns asynchronous job results containing temporary edit/view URLs. [Canva Design Import](https://www.canva.dev/docs/apps/rest-apis/reference/design-imports/create-design-import-job/)

**Not v1:** bidirectional synchronization. Canva can be used to continue editing a presentation, but Pekora must warn that post-handoff Canva edits will not update Pekora’s deck JSON.

**Experimental later:** Canva design generation. Canva can generate a presentation from a brief and slide outline, but this API is preview-only, consumes the user’s Canva AI credits, and should never replace Pekora’s deterministic deck engine. [Canva Design Generation](https://www.canva.dev/docs/apps/rest-apis/reference/generations/create-design-generation-job/)

---

## 3. Project structure

```text
pekora/
├── apps/
│   ├── api/
│   │   └── app/
│   │       ├── api/             # FastAPI routes
│   │       ├── contracts/       # Pydantic DTOs / JSON schemas
│   │       ├── core/            # config, errors, security
│   │       ├── db/              # SQLite models and migrations
│   │       ├── ingest/          # PDF, URL, OCR, source chunks
│   │       ├── ai/              # provider adapters, router, validator
│   │       ├── images/          # upload, search, approval, attribution
│   │       ├── builder/         # fixed layouts and python-pptx output
│   │       ├── export/          # PDF, PNG and ZIP creation
│   │       ├── scripts/         # speaker script generation
│   │       └── canva/           # OAuth, token store, import/generation client
│   └── web/
│       └── src/features/        # create, editor, images, export, canva
├── assets/fonts/
├── data/                        # ignored: SQLite, assets, cache, exports
├── docker-compose.yml
└── .env.example
```

---

## 4. Data contracts and database records

### Essential records

| Record | Purpose |
| --- | --- |
| `projects` | Brief, mode, grounding setting, selected theme, current status |
| `sources` | Normalized text chunks and provenance (file/page/URL/chunk) |
| `outlines` | Revisioned outline used before image search/deck generation |
| `decks` | Current validated deck JSON and revision number |
| `assets` | Uploaded/approved images, licensing, attribution, local cache path |
| `jobs` | Ingest, generation, render, script, Canva-import status/progress |
| `provider_usage` | Model/provider usage, categorized errors, circuit-breaker state |
| `canva_connections` | Encrypted Canva refresh token, granted scopes, connection status |
| `canva_handoffs` | Deck revision, export artifact hash, Canva import job/design ID, edit-link expiry |

Never store a Canva access token in the browser or in unencrypted SQLite fields.

### Slide schema

```python
# apps/api/app/contracts/deck.py
from enum import Enum
from uuid import UUID, uuid4
from pydantic import BaseModel, Field, model_validator

class LayoutId(str, Enum):
    TITLE = "title"
    BULLETS = "bullets"
    TWO_COLUMN = "two_column"
    IMAGE_TEXT = "image_text"
    QUOTE = "quote"
    CLOSING = "closing"

class AssetRef(BaseModel):
    asset_id: UUID
    role: str = "primary_image"

class Slide(BaseModel):
    id: UUID = Field(default_factory=uuid4)
    order: int = Field(ge=0)
    layout_id: LayoutId
    title: str = Field(min_length=1, max_length=120)
    bullets: list[str] = Field(default_factory=list, max_length=6)
    assets: list[AssetRef] = Field(default_factory=list, max_length=3)
    speaker_notes: str = ""
    source_refs: list[str] = Field(default_factory=list)
    verification_flags: list[str] = Field(default_factory=list)

    @model_validator(mode="after")
    def check_layout_assets(self):
        if self.layout_id == LayoutId.IMAGE_TEXT and not self.assets:
            raise ValueError("image_text requires an approved image asset")
        if self.layout_id != LayoutId.IMAGE_TEXT and self.assets:
            raise ValueError("this layout does not accept image assets")
        return self

class Deck(BaseModel):
    id: UUID = Field(default_factory=uuid4)
    project_id: UUID
    revision: int = Field(ge=1)
    slides: list[Slide] = Field(min_length=1, max_length=50)
```

The builder validates per-layout title/bullet limits separately from the model schema. That makes it possible to refine layout budgets without changing every API contract.

---

## 5. Main generation pipeline

```text
Input
  → ingest + provenance
  → project brief
  → validated outline
  → optional image search + explicit approval
  → validated slide JSON
  → limited corrections
  → PPTX build
  → PDF render
  → PNG ZIP
  → optional Canva import handoff
```

### Input rules

- A user-uploaded image must be classified as **slide image**, **source to read**, or **theme reference**.
- URL fetches reject local/private/reserved IP destinations, including redirects.
- PDF/OCR/URL content is wrapped as untrusted source data; it is never treated as model instructions.
- Grounded mode uses only input sources. Creative mode can add material but marks unsourced factual claims for review.

### Image approval gate

Pekora uses clear-license sources first (for example Unsplash, Pexels, Pixabay, Wikimedia Commons, Openverse, and eligible public collections). It displays candidate source/license data and thumbnails. The user must approve, reject, replace, or skip each candidate before deck generation. Any open-web source is disabled by default and labelled as license-unknown when enabled.

---

## 6. Deterministic PPTX, PDF, and PNG generation

```python
# apps/api/app/builder/service.py
from pathlib import Path
from pptx import Presentation
from app.contracts.deck import Deck

class DeckBuilder:
    def build(self, deck: Deck, output: Path) -> Path:
        ppt = Presentation()
        ppt.slide_width = 13_333_333  # 16:9, in EMU
        ppt.slide_height = 7_500_000

        for slide in deck.slides:
            self.validate_layout_budget(slide)
            self.add_fixed_layout(ppt, slide)

        ppt.save(output)
        return output

    def validate_layout_budget(self, slide) -> None:
        if len(slide.title) > 100:
            raise ValueError(f"Slide {slide.order + 1}: title is too long")
        if len(slide.bullets) > 5:
            raise ValueError(f"Slide {slide.order + 1}: too many bullets")
        if any(len(item) > 110 for item in slide.bullets):
            raise ValueError(f"Slide {slide.order + 1}: bullet is too long")

    def add_fixed_layout(self, ppt: Presentation, slide) -> None:
        # Dispatch by owned LayoutId. Never use AI-generated coordinates.
        raise NotImplementedError
```

```python
# apps/api/app/export/render.py
import subprocess
from pathlib import Path
import fitz  # PyMuPDF

def render_to_pdf_and_png(pptx: Path, work_dir: Path, dpi: int = 150):
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(work_dir), str(pptx)],
        cwd=work_dir, check=True, timeout=90,
    )
    pdf = work_dir / f"{pptx.stem}.pdf"
    document = fitz.open(pdf)
    images = []
    for i, page in enumerate(document):
        target = work_dir / f"slide-{i + 1:02d}.png"
        page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False).save(target)
        images.append(target)
    return pdf, images
```

Run LibreOffice only in a job-specific temporary directory. A release test gallery must render every layout, theme, font, and maximum-content fixture.

---

## 7. Multi-model routing and local mode

### Provider contract

```python
# apps/api/app/ai/base.py
from typing import Protocol
from app.contracts.deck import Deck

class GenerationError(Exception):
    def __init__(self, kind: str, message: str):
        self.kind = kind  # rate_limit, quota, timeout, invalid_output, auth, refusal
        super().__init__(message)

class DeckProvider(Protocol):
    name: str
    async def generate_deck(self, request) -> Deck: ...
    async def generate_script(self, request): ...
```

### Router policy

- Outline/slide JSON: reliable structured-output cloud model → alternate provider → OpenRouter model → detected Ollama → non-AI deck when appropriate.
- Speaker script: stronger writing provider → alternate provider → Ollama with quality warning.
- Use bounded JSON repair, max attempts, categorized errors, and circuit breakers.
- Display the actual model/provider used on the result.

### Ollama rule

Pekora calls a configurable local endpoint (normally `http://127.0.0.1:11434`) to list installed models. Local mode is unavailable if that check fails; Pekora must not download models. Manual RAM/VRAM/unified-memory settings are recommendations, not claimed hardware detection.

---

## 8. Canva implementation plan

### 8.1 Canva scope and user experience

Add Canva only after Pekora’s native PPTX/PDF/PNG route works.

1. User clicks **Connect Canva**.
2. Pekora opens the Canva OAuth consent page.
3. Canva redirects back to Pekora; Pekora securely stores the token connection.
4. User finishes/corrects a Pekora deck and selects **Open in Canva**.
5. Pekora builds the current PPTX revision and starts a Canva import job.
6. Pekora polls the job, stores the Canva design metadata, then presents **Continue editing in Canva**.
7. The UI states: “Changes made in Canva stay in Canva and do not update this Pekora deck.”

Canva REST APIs use OAuth 2.0 Authorization Code with PKCE; token exchange must occur server-side. [Canva OAuth documentation](https://www.canva.dev/docs/apps/rest-apis/authentication/)

### 8.2 Minimum OAuth scopes

| Feature | Scope |
| --- | --- |
| Import Pekora PPTX into Canva | `design:content:write` |
| Read/refresh design metadata and links | `design:meta:read` |
| Future asset upload | `asset:write` (and `asset:read` if reading assets) |
| Experimental Canva design generation | `design:content:write`, plus Canva’s user `design_generation` capability |

Request only features actually enabled. Canva scopes are explicit; a write scope does not imply its matching read scope. [Canva scope reference](https://www.canva.dev/docs/apps/rest-apis/scopes/)

### 8.3 OAuth route and PKCE state

```python
# apps/api/app/canva/oauth.py
import base64
import hashlib
import secrets
from urllib.parse import urlencode

CANVA_AUTHORIZE_URL = "https://www.canva.com/api/oauth/authorize"

def begin_authorization(client_id: str, redirect_uri: str) -> tuple[str, dict]:
    verifier = secrets.token_urlsafe(72)[:128]
    challenge = base64.urlsafe_b64encode(
        hashlib.sha256(verifier.encode()).digest()
    ).rstrip(b"=").decode()
    state = secrets.token_urlsafe(32)
    params = {
        "code_challenge": challenge,
        "code_challenge_method": "S256",
        "scope": "design:content:write design:meta:read",
        "response_type": "code",
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "state": state,
    }
    # Persist state + verifier server-side with a short expiration before redirecting.
    return f"{CANVA_AUTHORIZE_URL}?{urlencode(params)}", {"state": state, "verifier": verifier}
```

```python
# apps/api/app/api/routes/canva.py
from fastapi import APIRouter, HTTPException

router = APIRouter(prefix="/integrations/canva", tags=["canva"])

@router.get("/callback")
async def canva_callback(code: str, state: str):
    pending = await oauth_state_store.consume(state)
    if pending is None:
        raise HTTPException(status_code=400, detail="Invalid or expired Canva authorization state")
    # Exchange `code` server-side using pending.verifier and the client secret.
    # Encrypt refresh/access credentials before persistence.
    await canva_service.complete_authorization(code, pending.verifier)
    return {"connected": True}
```

Never expose `client_secret`, authorization-code exchange, `code_verifier`, refresh token, or Canva access token to React.

### 8.4 Import the current Pekora PPTX

Canva imports external files through `POST /rest/v1/imports` using the raw file bytes, `application/octet-stream`, and an `Import-Metadata` header. Its import result is asynchronous. [Canva import endpoint](https://www.canva.dev/docs/apps/rest-apis/reference/design-imports/create-design-import-job/)

```python
# apps/api/app/canva/client.py
import base64
import json
from pathlib import Path
import httpx

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

class CanvaClient:
    def __init__(self, access_token: str):
        self.headers = {"Authorization": f"Bearer {access_token}"}

    async def import_pptx(self, pptx_path: Path, title: str) -> str:
        metadata = {
            "title_base64": base64.b64encode(title[:50].encode("utf-8")).decode(),
            "mime_type": PPTX_MIME,
        }
        headers = {
            **self.headers,
            "Content-Type": "application/octet-stream",
            "Import-Metadata": json.dumps(metadata),
        }
        async with httpx.AsyncClient(timeout=60) as client:
            response = await client.post(
                "https://api.canva.com/rest/v1/imports",
                headers=headers,
                content=pptx_path.read_bytes(),
            )
            response.raise_for_status()
        return response.json()["job"]["id"]

    async def get_import_job(self, job_id: str) -> dict:
        async with httpx.AsyncClient(timeout=15) as client:
            response = await client.get(
                f"https://api.canva.com/rest/v1/imports/{job_id}",
                headers=self.headers,
            )
            response.raise_for_status()
        return response.json()["job"]
```

Production refinement: stream the file instead of `read_bytes()` for larger exports, refresh an expired access token before calls, apply a bounded polling/backoff policy, and map Canva errors (`invalid_file`, throttling, duplicate import) to user-friendly messages.

### 8.5 Pekora handoff route

```python
# apps/api/app/api/routes/canva_handoff.py
@router.post("/decks/{deck_id}/canva-handoff", status_code=202)
async def create_canva_handoff(deck_id: str):
    connection = await canva_connections.require_active()
    deck = await deck_repository.get_current(deck_id)
    pptx = await export_service.build_current_pptx(deck)
    job_id = await CanvaClient(await connection.access_token()).import_pptx(pptx, deck.title)
    job = await jobs.create(kind="canva_import", external_id=job_id, deck_id=deck.id)
    return {"job_id": str(job.id), "status": "queued"}
```

When the import job succeeds, record Canva’s design ID and its `edit_url`. Canva documents these returned URLs as temporary (currently 30 days), so Pekora must refresh design information rather than treating an old saved link as permanent. [Canva import-job response](https://www.canva.dev/docs/apps/rest-apis/reference/design-imports/create-design-import-job/)

### 8.6 React handoff UI

```tsx
// apps/web/src/features/canva/CanvaHandoffButton.tsx
export function CanvaHandoffButton({ deckId, connected }: { deckId: string; connected: boolean }) {
  const [busy, setBusy] = useState(false);

  async function handoff() {
    setBusy(true);
    try {
      const response = await api.post(`/decks/${deckId}/canva-handoff`);
      await waitForJob(response.data.job_id);
      // Job result contains a temporary Canva edit URL.
      window.open((await api.get(`/jobs/${response.data.job_id}`)).data.result.edit_url, "_blank", "noopener,noreferrer");
    } finally {
      setBusy(false);
    }
  }

  if (!connected) return <a href="/api/integrations/canva/authorize">Connect Canva</a>;
  return <button disabled={busy} onClick={handoff}>{busy ? "Sending to Canva…" : "Open in Canva"}</button>;
}
```

Place the one-way-sync notice immediately beside this action, not in a hidden settings page.

---

## 9. Optional Canva AI mode — after MVP only

This is a separate experiment, not a model-provider fallback.

- User explicitly selects **Generate design in Canva (Experimental)**.
- Pekora sends a compact approved brief and outline, not unreviewed full source text by default.
- Check the user’s `design_generation` capability before enabling the action.
- Handle `credit_quota_exceeded`, `model_not_available`, safety rejection, and generation failure without silently using another service.
- Save the returned Canva design link as a Canva-owned result; do not attempt to reconstruct it as Pekora slide JSON.

Canva states this API is preview-only and public apps using preview APIs cannot pass review; its use should therefore remain personal/experimental until Canva’s status changes. [Canva Generations API](https://www.canva.dev/docs/apps/rest-apis/reference/generations/)

---

## 10. Delivery phases and completion checks

| Phase | Deliverable | Done when |
| --- | --- | --- |
| 0 | Repository, contracts, fonts, local tools | Hand-authored JSON validates and renders locally. |
| 1 | Fixed layouts, PPTX/PDF/PNG | Layout gallery exports correct page/slide/image counts. |
| 2 | Projects, jobs, quick-create UI, non-AI deck | Topic to text-only export works without a model. |
| 3 | First cloud model + validator/router | Invalid model output cannot reach the builder; fallback is bounded. |
| 4 | Themes, PDF/URL/OCR ingest, provenance | Grounded deck preserves source references and rejects unsafe URLs. |
| 5 | Image approval gallery | No unapproved candidate can appear in exported deck JSON or PPTX. |
| 6 | Corrective editor + speaker script | Corrections revise exports; script becomes stale after relevant edit. |
| 7 | Ollama + additional model providers | Local option appears only when a compatible installed model is detected. |
| 8 | Canva Connect + PPTX import handoff | User authorizes Canva, imports current deck, opens returned edit URL, and sees one-way-sync warning. |
| 9 | Experimental Canva AI (optional) | Only opt-in users with capability/credits can submit; failures are clear and isolated. |

---

## 11. Test plan

### Native Pekora tests

- Schema/layout/overflow validation.
- PPTX → PDF → PNG gallery tests across every theme/layout.
- Model provider mock failures: timeout, quota, 429, malformed JSON, refusal.
- URL SSRF/redirect checks; PDF/image file validation; prompt-injection fixtures.
- Image candidate approval enforcement, duplicate removal, attribution preservation.
- Ollama absent/present/oversized-model states.

### Canva-specific tests

- PKCE state is single-use, expiring, and verified before token exchange.
- OAuth client secret/token data never appear in logs, API responses, or browser storage.
- Unconnected user cannot request a handoff.
- A deck handoff always exports the exact current deck revision and stores that revision ID.
- Import `in_progress`, `success`, `failed`, throttled, invalid-file, and expired-token cases update Pekora jobs correctly.
- Returned edit links open only after job success; UI labels them temporary and shows the one-way-sync notice.
- Canva design-generation UI remains disabled without user capability and is excluded from normal Pekora provider fallback tests.

---

## 12. Configuration and deployment notes

```text
APP_DATA_DIR=./data
DATABASE_URL=sqlite:///./data/pekora.db
LIBREOFFICE_BIN=soffice
OLLAMA_BASE_URL=http://127.0.0.1:11434

CANVA_CLIENT_ID=...
CANVA_CLIENT_SECRET=...
CANVA_REDIRECT_URI=http://localhost:8000/api/integrations/canva/callback
TOKEN_ENCRYPTION_KEY=...

GEMINI_API_KEY=...
GROQ_API_KEY=...
OPENROUTER_API_KEY=...
```

For personal local use, run the API and UI on the same computer as LibreOffice and Ollama. A remote deployment can support cloud models and Canva but cannot assume access to local Ollama without an explicit, secure local-agent/network design.

Before public deployment, add authentication, CSRF protection, secure token key management, upload limits, retention/deletion rules, a reverse proxy, backups, and a formal privacy review.

---

## 13. Version-one boundary

Pekora v1 is complete when a user can create, edit lightly, and export a grounded or creative presentation; approve images before use; generate a script; use cloud or local generation when available; and hand off the final PPTX to Canva for optional further editing.

Pekora does **not** promise bidirectional Canva sync, arbitrary free-form editing, generated images, multi-user collaboration, automatic model installation, or exact visual equality between LibreOffice and PowerPoint.
