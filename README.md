# Pekora — Implementation Plan

> **Project name:** Pekora  
> **Scope:** Personal, local-first PPT generator web app.  
> **Status:** Build plan with representative starter code. Snippets establish interfaces and patterns; they are not a complete production application.

---

## 1. Product contract

Pekora accepts a topic, pasted text, URLs, PDFs, and images; creates a presentation using fixed layouts; permits small corrections; and exports PPTX, PDF, and PNG slides. It optionally generates a presentation script.

Non-negotiable rules:

1. Slide JSON is the source of truth; never edit a generated PPTX as application state.
2. Models provide validated content, not slide geometry or templates.
3. Internet images are searched only when requested and are included only after explicit approval.
4. Ollama is optional and shown only when a local service and installed model are detected.
5. A non-AI fallback creates a plain deck from source headings/paragraphs.

---

## 2. Proposed stack

| Layer | Technology | Why |
| --- | --- | --- |
| API | Python 3.11+, FastAPI, Pydantic v2 | Strong schemas and async APIs |
| UI | React, TypeScript, Vite, PWA | Fast local UI and later mobile-friendly installation |
| Storage | SQLite + SQLAlchemy/Alembic | Appropriate for one local user |
| PPTX | python-pptx | Deterministic PowerPoint creation |
| PDF/PNG | LibreOffice headless + PyMuPDF | One PPTX-based export path |
| Ingest | PyMuPDF, Tesseract, trafilatura, Pillow | PDF, OCR, URL, and image processing |
| AI | Gemini, Groq, OpenRouter, one direct paid provider, Ollama | Task-based routing and fallback |
| Jobs | SQLite job records + in-process worker initially | Observable local jobs without premature infrastructure |

Use Docker Compose only after the local developer path works; it is a packaging option, not a prerequisite for the first vertical slice.

---

## 3. Repository shape

```text
pekora/
├── apps/
│   ├── api/
│   │   ├── app/
│   │   │   ├── main.py
│   │   │   ├── core/          # settings, errors, security
│   │   │   ├── contracts/     # Pydantic schemas
│   │   │   ├── api/           # FastAPI routers
│   │   │   ├── services/      # orchestration and use cases
│   │   │   ├── ai/            # provider adapters and router
│   │   │   ├── builder/       # layouts and PPTX build
│   │   │   ├── ingest/        # PDF, URL, OCR
│   │   │   ├── images/        # assets, searches, approval
│   │   │   ├── export/        # LibreOffice/PDF/PNG
│   │   │   └── db/            # persistence
│   │   └── tests/
│   └── web/
│       └── src/features/      # create, images, editor, export
├── packages/contracts/        # optional generated TS types/JSON schemas
├── assets/fonts/
├── data/                      # ignored: db, uploads, cache, exports
├── docker-compose.yml
└── .env.example
```

---

## 4. Phase 0 — Establish the contracts before integrations

Create the shared objects first. Every provider, editor action, export, and test consumes these structures.

### 4.1 Pydantic slide contract

```python
# apps/api/app/contracts/deck.py
from enum import Enum
from typing import Literal
from uuid import UUID, uuid4
from pydantic import BaseModel, Field, HttpUrl, model_validator


class LayoutId(str, Enum):
    TITLE = "title"
    BULLETS = "bullets"
    TWO_COLUMN = "two_column"
    IMAGE_TEXT = "image_text"
    QUOTE = "quote"
    CLOSING = "closing"


class SourceRef(BaseModel):
    source_id: UUID
    locator: str | None = None  # e.g. PDF page, URL fragment, chunk id


class AssetRef(BaseModel):
    asset_id: UUID
    role: Literal["primary_image", "supporting_image"]


class TextBlock(BaseModel):
    kind: Literal["bullets", "paragraph", "quote", "caption"]
    text: str | None = None
    bullets: list[str] = Field(default_factory=list, max_length=6)


class Slide(BaseModel):
    id: UUID = Field(default_factory=uuid4)
    order: int = Field(ge=0)
    layout_id: LayoutId
    title: str = Field(min_length=1, max_length=120)
    blocks: list[TextBlock] = Field(default_factory=list, max_length=3)
    assets: list[AssetRef] = Field(default_factory=list, max_length=3)
    speaker_notes: str = ""
    source_refs: list[SourceRef] = Field(default_factory=list)
    verification_flags: list[str] = Field(default_factory=list)
    alt_text: str | None = None

    @model_validator(mode="after")
    def layout_rules(self) -> "Slide":
        if self.layout_id == LayoutId.IMAGE_TEXT and not self.assets:
            raise ValueError("image_text slides require an approved image asset")
        if self.layout_id != LayoutId.IMAGE_TEXT and len(self.assets) > 0:
            raise ValueError("this initial layout catalog does not accept images here")
        return self


class Deck(BaseModel):
    id: UUID = Field(default_factory=uuid4)
    project_id: UUID
    theme_id: UUID
    slides: list[Slide] = Field(min_length=1, max_length=50)

    @model_validator(mode="after")
    def slide_order_is_contiguous(self) -> "Deck":
        if [s.order for s in self.slides] != list(range(len(self.slides))):
            raise ValueError("slides must be ordered contiguously from zero")
        return self
```

Keep layout-specific text budgets in a separate catalog rather than scattering limits in prompt code:

```python
# apps/api/app/builder/layouts.py
from dataclasses import dataclass
from app.contracts.deck import LayoutId

@dataclass(frozen=True)
class LayoutSpec:
    max_title_chars: int
    max_bullets: int
    max_bullet_chars: int
    accepts_images: bool

LAYOUTS: dict[LayoutId, LayoutSpec] = {
    LayoutId.TITLE: LayoutSpec(70, 0, 0, False),
    LayoutId.BULLETS: LayoutSpec(100, 5, 110, False),
    LayoutId.TWO_COLUMN: LayoutSpec(90, 6, 75, False),
    LayoutId.IMAGE_TEXT: LayoutSpec(90, 4, 90, True),
    LayoutId.QUOTE: LayoutSpec(80, 0, 0, False),
    LayoutId.CLOSING: LayoutSpec(90, 4, 90, False),
}
```

**Done when:** invalid layout/asset combinations and overflow-risk payloads cannot enter the builder.

---

## 5. Phase 1 — Build the deterministic deck engine

Implement a hand-authored deck fixture → PPTX → PDF → PNG route before adding AI.

### 5.1 Builder boundary

```python
# apps/api/app/builder/service.py
from pathlib import Path
from pptx import Presentation
from app.contracts.deck import Deck
from app.builder.layouts import LAYOUTS


class DeckBuilder:
    def build(self, deck: Deck, output_path: Path) -> Path:
        presentation = Presentation()
        presentation.slide_width = 13_333_333   # 16:9 EMU
        presentation.slide_height = 7_500_000

        for slide in deck.slides:
            self._validate_budget(slide)
            self._add_slide(presentation, slide, deck.theme_id)

        presentation.save(output_path)
        return output_path

    def _validate_budget(self, slide) -> None:
        spec = LAYOUTS[slide.layout_id]
        if len(slide.title) > spec.max_title_chars:
            raise ValueError(f"Slide {slide.id}: title exceeds layout budget")
        for block in slide.blocks:
            if len(block.bullets) > spec.max_bullets:
                raise ValueError(f"Slide {slide.id}: too many bullets")
            if any(len(item) > spec.max_bullet_chars for item in block.bullets):
                raise ValueError(f"Slide {slide.id}: bullet exceeds layout budget")

    def _add_slide(self, presentation: Presentation, slide, theme_id) -> None:
        # Dispatch only to owned, tested layout functions.
        # Never render arbitrary coordinates returned by an AI model.
        raise NotImplementedError
```

### 5.2 Export pipeline

```python
# apps/api/app/export/render.py
import subprocess
from pathlib import Path
import fitz  # PyMuPDF


def render_pptx(pptx_path: Path, work_dir: Path, png_dpi: int = 150) -> tuple[Path, list[Path]]:
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(work_dir), str(pptx_path)],
        check=True,
        timeout=90,
        cwd=work_dir,
    )
    pdf_path = work_dir / f"{pptx_path.stem}.pdf"
    if not pdf_path.exists():
        raise RuntimeError("LibreOffice did not produce PDF output")

    document = fitz.open(pdf_path)
    pngs: list[Path] = []
    zoom = png_dpi / 72
    for index, page in enumerate(document):
        pixmap = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        target = work_dir / f"slide-{index + 1:02d}.png"
        pixmap.save(target)
        pngs.append(target)
    return pdf_path, pngs
```

Run LibreOffice inside a job-specific temporary directory. Add process limits and cleanup policy before exposing the endpoint beyond local use.

**Done when:** a six-slide fixture produces a PPTX, six-page PDF, and ZIP containing six PNGs.

---

## 6. Phase 2 — API and project/job persistence

Use explicit job records even if the first worker is in-process. That makes progress, retries, and future queue migration possible.

### 6.1 API shape

| Endpoint | Function |
| --- | --- |
| `POST /projects` | Create draft project/brief |
| `POST /projects/{id}/ingest` | Add text, file, or URL sources |
| `POST /projects/{id}/outline` | Queue outline generation |
| `POST /projects/{id}/images/search` | Search images for an outline revision |
| `POST /projects/{id}/images/approve` | Approve/reject a candidate asset |
| `POST /projects/{id}/generate` | Generate validated deck JSON |
| `PATCH /decks/{id}/slides/{slide_id}` | Perform a corrective edit |
| `POST /decks/{id}/exports` | Queue PPTX/PDF/PNG export |
| `POST /decks/{id}/script` | Generate speaker script |
| `GET /jobs/{id}` | Read state/progress/result |
| `GET /providers/status` | Read configured/detected capabilities only |

### 6.2 Example route and use case

```python
# apps/api/app/api/routes/projects.py
from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel, Field
from app.services.projects import ProjectService

router = APIRouter(prefix="/projects", tags=["projects"])

class CreateProjectRequest(BaseModel):
    title: str = Field(min_length=1, max_length=160)
    topic: str | None = Field(default=None, max_length=8_000)
    slide_count: int = Field(default=8, ge=1, le=50)
    mode: str = Field(default="quick", pattern="^(quick|guided)$")

@router.post("", status_code=status.HTTP_201_CREATED)
async def create_project(payload: CreateProjectRequest, service: ProjectService) -> dict:
    project = await service.create(payload)
    return {"id": str(project.id), "status": project.status}
```

The actual dependency provider for `ProjectService`, database sessions, authentication (if later needed), and exception mapping belongs in `main.py`/dependency modules—not inside routes.

**Done when:** a browser can create a project, observe a job, and retrieve the current deck JSON.

---

## 7. Phase 3 — Ingestion and safety boundary

Input sources are data, never instructions. Preserve provenance so grounded decks and scripts can point back to pages/chunks.

```python
# apps/api/app/ingest/url_guard.py
import ipaddress
import socket
from urllib.parse import urlparse

BLOCKED_HOSTS = {"localhost", "localhost.localdomain"}

def assert_public_http_url(url: str) -> None:
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise ValueError("Only absolute HTTP(S) URLs are accepted")
    if parsed.hostname.lower() in BLOCKED_HOSTS:
        raise ValueError("Local URLs are not allowed")

    addresses = socket.getaddrinfo(parsed.hostname, None, type=socket.SOCK_STREAM)
    for _, _, _, _, address in addresses:
        ip = ipaddress.ip_address(address[0])
        if not ip.is_global:
            raise ValueError("Non-public network addresses are not allowed")
```

Re-run this check for each redirect target, impose small response-size/time limits, and verify file signatures after download. Do not rely only on this sample code for production SSRF controls; resolver rebinding and proxy/network policy need integration tests.

**Done when:** PDF pages, URL text, and OCR output become normalized chunks with source IDs; local/private URLs and oversize/spoofed files are rejected safely.

---

## 8. Phase 4 — Provider abstraction and safe fallback

Models return contracts, never arbitrary markdown that the builder tries to interpret.

```python
# apps/api/app/ai/providers/base.py
from typing import Protocol
from app.contracts.deck import Deck

class GenerationError(Exception):
    def __init__(self, kind: str, message: str):
        self.kind = kind  # rate_limit, quota, timeout, invalid_output, auth, refusal
        super().__init__(message)

class DeckProvider(Protocol):
    name: str

    async def generate_deck(self, request: "DeckGenerationRequest") -> Deck:
        ...

    async def generate_script(self, request: "ScriptRequest") -> "SpeakerScript":
        ...
```

```python
# apps/api/app/ai/router.py
class ProviderRouter:
    def __init__(self, providers: list[DeckProvider], circuit_breaker, usage_store):
        self.providers = providers
        self.circuit_breaker = circuit_breaker
        self.usage_store = usage_store

    async def generate_deck(self, request):
        failures = []
        for provider in self.providers:
            if self.circuit_breaker.is_open(provider.name):
                continue
            try:
                result = await provider.generate_deck(request)
                self.usage_store.record_success(provider.name)
                return result
            except GenerationError as exc:
                self.usage_store.record_failure(provider.name, exc.kind)
                self.circuit_breaker.record_failure(provider.name, exc.kind)
                failures.append({"provider": provider.name, "kind": exc.kind})
        raise RuntimeError({"message": "No configured model could generate a deck", "failures": failures})
```

Use distinct sequences per task. Outline and schema work prioritize structured-output reliability; speaker scripts prioritize writing quality. Include Ollama only in the sequence when its health endpoint confirms a selected installed model.

**Done when:** mocked 429, timeout, and malformed JSON responses select the next eligible provider, stop within the retry budget, and expose a useful final failure.

---

## 9. Phase 5 — Image search and explicit approval

Search happens after outline generation and before deck generation. The deck generator receives only approved `asset_id` references.

```python
# apps/api/app/images/service.py
from uuid import UUID

class ImageService:
    async def approve(self, project_id: UUID, candidate_id: UUID) -> UUID:
        candidate = await self.repository.get_candidate(project_id, candidate_id)
        if candidate is None:
            raise ValueError("Unknown image candidate")
        local_asset = await self.downloader.fetch_and_normalize(candidate.url)
        return await self.repository.create_asset(
            project_id=project_id,
            path=local_asset.path,
            sha256=local_asset.sha256,
            source_url=candidate.url,
            license_name=candidate.license_name,
            attribution=candidate.attribution,
            approved=True,
        )
```

Required implementation details:

- Search clear-license adapters in parallel; make open-web search an off-by-default capability.
- Normalize license/source/author/width/height metadata before ranking.
- Apply perceptual deduplication, resolution/aspect filtering, and source-level timeouts.
- Show source and license badges in the UI, with approve/reject/search again/upload/skip controls.
- Revalidate direct image URLs using the same URL guard; strip EXIF, validate MIME/signature, resize, convert to PNG/JPEG, and cache locally.

**Done when:** an image cannot be attached to a slide unless an approved asset record exists.

---

## 10. Phase 6 — React flow and corrective editor

Keep client state small. The server is authoritative for persisted deck/job state.

```tsx
// apps/web/src/features/editor/SlideEditor.tsx
type SlideEditorProps = {
  slide: Slide;
  onSave: (patch: Pick<Slide, "title" | "blocks" | "speaker_notes" | "layout_id">) => Promise<void>;
};

export function SlideEditor({ slide, onSave }: SlideEditorProps) {
  const [title, setTitle] = useState(slide.title);
  const [notes, setNotes] = useState(slide.speaker_notes);

  return (
    <section aria-label={`Edit slide ${slide.order + 1}`}>
      <input value={title} onChange={(event) => setTitle(event.target.value)} />
      <textarea value={notes} onChange={(event) => setNotes(event.target.value)} />
      <button onClick={() => onSave({ title, blocks: slide.blocks, speaker_notes: notes, layout_id: slide.layout_id })}>
        Save correction
      </button>
    </section>
  );
}
```

The patch endpoint must re-run schema/layout/asset validation and return the updated slide plus overflow warnings. Support undo/redo with versioned deck revisions or an append-only local edit history—choose one during implementation, but never rely solely on browser history.

**Done when:** title/text/notes/layout corrections update the server deck and rebuild all exports; free-form object editing remains unavailable by design.

---

## 11. Phase 7 — Local mode and speaker scripts

### Ollama detection

```python
# apps/api/app/ai/ollama.py
import httpx

async def get_ollama_status(base_url: str) -> dict:
    try:
        async with httpx.AsyncClient(timeout=2.0) as client:
            response = await client.get(f"{base_url.rstrip('/')}/api/tags")
            response.raise_for_status()
    except httpx.HTTPError:
        return {"available": False, "models": []}
    models = [item["name"] for item in response.json().get("models", [])]
    return {"available": True, "models": models}
```

Never auto-download models. Manual RAM/VRAM/unified-memory selections offer guidance, but they do not claim to detect hardware. Local generation uses outline-first, slide-by-slide structured output with an explicit context window. Offline with no Ollama still exposes the non-AI builder.

### Script rules

Generate scripts from the final deck revision. Mark a script stale when a slide changes. In Grounded mode, only include sourced slide facts plus non-factual transitions/questions/analogies. In Creative mode, flag unsourced claims for review.

**Done when:** the UI accurately reports local capability, never tries unavailable Ollama, and a script is revision-linked to the deck.

---

## 12. Verification plan

| Suite | Essential assertions |
| --- | --- |
| Contracts | Invalid orders, overflow content, illegal image/layout pair, and missing asset approval fail validation. |
| Builder/render | Every layout fixture creates PPTX, matching PDF pages, and expected PNG count. |
| Providers | 429, quota, timeout, refusal, and malformed JSON result in bounded fallback/circuit-break behavior. |
| Ingest/security | Private IP, redirect-to-private-IP, giant response, spoofed MIME, malicious PDF, and injection text are safely handled. |
| Images | Unapproved, duplicate, low-resolution, failed-download, and license-unknown candidates behave as specified. |
| Local mode | Ollama unavailable/available/oversized selection states are clear; no-Ollama offline fallback works. |
| E2E | Topic → outline → optional image approval → deck → correction → PPTX/PDF/PNG/script. |

Use a rendered “layout gallery” fixture as a release gate. It should include longest permitted titles/bullets, representative images, every bundled font, and every theme.

---

## 13. Build sequence

1. **Week 1:** contracts, SQLite models, fixture deck, one deterministic layout.
2. **Week 2:** remaining fixed layouts, PPTX/PDF/PNG pipeline, layout-gallery test.
3. **Week 3:** project/job API and React quick-create flow, non-AI deck.
4. **Weeks 4–5:** first cloud provider, strict validation, provider failures, progress UI.
5. **Weeks 6–7:** themes, PDF/URL/OCR ingestion, source provenance and safety tests.
6. **Weeks 8–9:** image search/approval/cache/attribution.
7. **Week 10:** minimal editor and script generation.
8. **Weeks 11–12:** additional providers, OpenRouter, Ollama, hardening, PWA and release documentation.

The timeline is an estimate for part-time solo work. Each phase ends with a small working release; do not begin a later phase until its acceptance criteria pass.

---

## 14. Explicit non-goals for version one

- Free-form PowerPoint-quality editing in the browser.
- Generated images.
- Multi-user collaboration, accounts, billing, or public SaaS hosting.
- Automatic Ollama installation/model downloads.
- Image inclusion without user approval.
- Guaranteeing perfect visual parity between LibreOffice and Microsoft PowerPoint.

These boundaries preserve the core promise: a reliable, editable-enough presentation generator whose output and behavior remain predictable across multiple AI providers.
