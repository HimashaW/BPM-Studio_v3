# BPM Studio - GuideX (Service Delivery Guide Generator)

BPM Studio is the internal BPM automation platform; GuideX is its first tool. `app.py` is the web UI; `app_service.py` holds the generation logic.
Users upload the workbook, fill in the document details and download: the Word guide, the swimlane
flowchart images and the editable PowerPoint (plus a ZIP of everything).

## BPM Studio platform layout (v1.1)

BPM Studio is now a small platform that hosts several tools. `app.py` is only the shell (sidebar, header,
Home page, router). Each tool lives in its own folder and is listed in `tools/__init__.py`.

```
app.py                  platform shell - does not change when a tool is added
studio/                 shared theme (CSS, logo), session-state helpers, ToolSpec type, stepper
tools/__init__.py       TOOLS = [guidex.SPEC, testx.SPEC]   <- the one line to add for a new tool
tools/guidex/           GuideX screens (page.py) + icon; generation code is still app_service.py / pipeline_core.py
tools/testx/            TestX - Test Case Generator
    page.py             the four screens
    source.py           reads the workbook with the existing GuideX extractors
    rules.py            business rules -> constraints (tables, tick matrices, 4 sentence forms)
    engine.py           pairwise / full combinations with rule constraints (no Streamlit)
    plan.py             picks attributes, runs the engine, labels positive / negative tests
    export.py           Excel UAT test plan
    mascot.py           Tess, the original TestX mascot (inline SVG)
tests/test_testx_engine.py   engine, rule, plan and Excel tests (fixture built by tests/fixtures/build_testx_sample.py)
```

### Adding another tool
1. Create `tools/<name>/` with `__init__.py` exposing `SPEC = ToolSpec(...)` (see `studio/registry.py`) and a `render()` function.
2. Prefix every session-state key with the tool's short id (`gx_`, `tx_`) so tools never clash.
3. Add the tool to `TOOLS` in `tools/__init__.py`. The sidebar, Home card, Help text and router appear automatically.
4. Keep logic out of the page file (like `engine.py` / `app_service.py`) so it can be tested without Streamlit.

## Hosting BPM Studio

**Run locally**
```bash
pip install -r requirements.txt
streamlit run app.py          # http://localhost:8501
```

**Option A - Streamlit Community Cloud (quickest, free)**
1. Push this folder to a GitHub repo (private).
2. share.streamlit.io -> New app -> pick the repo, branch `main`, main file `app.py`.
3. (Optional) Advanced settings -> Secrets: `GEMINI_API_KEY = "..."` so users need no key for the AI introduction.
4. Share the URL. Pushing to `main` redeploys automatically.

**Option B - Docker on your own server / cloud (recommended for company data)**
```bash
docker build -t bpm-studio .
docker run -p 8501:8501 bpm-studio
```
Works on any Docker host (Azure, AWS, Google Cloud Run, Render, Railway, or an internal VM).
Put it behind your company SSO / reverse proxy if only staff should reach it.

**Notes**
- Uploaded workbooks are processed in a temp folder and deleted after each run; nothing is stored.
- Product configuration workbooks can be confidential: prefer Option B on an internal network.
- Platform name/tagline constants are at the top of `app.py`; flow rules live in `mapping/flow_rules.yaml`.
- Upload limit is 25 MB (`.streamlit/config.toml`).

---

# BPM Service Delivery Guide Automation Pipeline

Generates a product Service Delivery Guide (.docx) from a config Excel
workbook, automatically. Fully product-agnostic — pass any product/service
name via `--product` and the title page, running header, and every section
heading retitle themselves accordingly.

## Quick start

```bash
pip install -r requirements.txt
python run.py --input input/Book1.xlsx --output output/ServiceDeliveryGuide.docx --product "Your Product Name"
python tests/test_extractors.py
```

## Web app (for day-to-day team use)

A Streamlit web UI (`app.py`) wraps the same pipeline behind an
upload-and-download page: pick an Excel file, fill in the document
metadata, click Generate, download the `.docx`. It uses the exact same
extraction/rendering code as the CLI (both call into `pipeline_core.py`),
so there's nothing to keep in sync between the two.

**Run it locally first, to confirm it works before deploying:**
```bash
pip install -r requirements.txt
streamlit run app.py
```
This opens the app at `http://localhost:8501`.

**Before deploying, add your real company logo:**
The web app does NOT ask each user to upload a logo - it's baked into
the deployment so every generated document is consistent. Just replace
the file at `assets/company_logo.png` with your actual logo (same exact
filename). If that file isn't present, the app automatically falls back
to the built-in gray placeholder and shows a warning after generation -
so nothing breaks if you forget this step, but the doc will have the
placeholder until you do it.

**Word template is the document master (default):**

The pipeline now uses `assets/process_guide_template.docx` as the default
Word document master. The supplied template is preserved as-is for all common
formatting and layout, including:

- cover page and company logo
- running header/footer
- page numbering
- Document Control Sheet layout
- Prepared / Checked / Approved tables
- Revision History table
- Contents page layout
- the `01. Introduction` section anchor
- fonts, colours, margins, spacing, and Word styles

The renderer **does not rebuild those elements in Python** when the template
is used. It only fills metadata/control values and inserts the dynamic content
extracted from Excel/YAML after the template's Introduction anchor.

The supplied template does not need `{{TOKENS}}`. Existing labels such as
`Title`, `Document Number`, `Issue No`, `Issue Date`, `Issued By`, and the
Document Control contact fields are populated directly while preserving the
Word formatting. Optional `{{PRODUCT_NAME}}`, `{{DOCUMENT_NO}}`,
`{{ISSUE_NO}}`, `{{ISSUE_DATE}}`, `{{REVISION_NO}}`, `{{REVISION_DATE}}`, and
`{{ISSUED_BY}}` tokens are also supported for future template versions.

To change the document format in the future, replace
`assets/process_guide_template.docx` with the new approved Word format. The
Excel extraction logic does not need to change as long as the template keeps
its `Introduction` anchor and the standard metadata labels, or uses the
supported `{{TOKEN}}` placeholders.

**Template override:**

```bash
python run.py \\
  --input input/Book1.xlsx \\
  --output output/ServiceDeliveryGuide.docx \\
  --product "Your Product" \\
  --template path/to/another_template.docx
```

If no `--template` is supplied, the bundled `assets/process_guide_template.docx`
is used automatically. If that file is unavailable, the legacy code-generated
layout remains as a safe fallback.

## Deploying for your team — Streamlit Community Cloud (recommended)

Free, gives your team one shared link, no server to maintain.

1. Push this project to a **private** GitHub repo (private matters here,
   since the code + any committed sample data shouldn't be public).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub, click "New app".
3. Point it at your repo, branch `main`, main file path `app.py`.
4. Deploy. You'll get a URL like `https://your-app-name.streamlit.app` —
   share that with your team.
5. Anyone with the link uploads their Excel, fills the form, downloads
   the generated `.docx`. Nothing is stored server-side between
   sessions — each run uses a temp directory that's discarded after.

**Updating it later:** just push to the repo's `main` branch; Streamlit
Community Cloud redeploys automatically within a minute or two.

**Access control:** Community Cloud apps are public by URL by default
(anyone with the link can use it, though the link itself isn't
discoverable/indexed). If your team needs actual login-gated access,
either keep the link private within the team, or ask about "Community
Cloud for Teams" / migrate to a platform with auth support (Render,
Railway, or an internal server — see below) if that's a hard
requirement.

### Alternative: Render / Railway (if you need more control)

Both support a `Dockerfile` or a native Python buildpack and give you a
persistent URL, environment variables, and optionally basic auth in
front of the app - useful if Streamlit Community Cloud's public-link
model isn't restrictive enough. Deployment is the same `streamlit run
app.py` command, just behind their build/start config instead of
Community Cloud's automatic detection. Both have free tiers sufficient
for a small team's occasional use.

## Full CLI example

```bash
python run.py \
  --input input/Book1.xlsx \
  --output output/ServiceDeliveryGuide.docx \
  --product "CYS_EdgeDefend" \
  --sheet "6.3 OSS SO Task List " \
  --doc-no "DOC-CYS-001" \
  --issue-no "001" \
  --revision-no "0" \
  --issued-by "BPM Section" \
  --logo "assets/your_logo.png" \
  --contact-name "J. Perera" \
  --contact-designation "DGM/BPM" \
  --contact-telephone "+94 11 1234567" \
  --contact-email "jperera@company.com"
```
Every flag except `--input`/`--output`/`--product` is optional with a
sensible default (see `run.py`'s `argparse` section for the full list).


### Generated section numbering

When the Word template is used, generated section numbers are written explicitly by the renderer rather than relying on Word's automatic multilevel numbering. This prevents the template's existing sample numbering from continuing into generated content. The workflow section structure is:

```text
4. Workflow Diagram - CREATE Order Type
4.1 Product Name
4.1.1 Task Flow - Product Name

5. Workflow Diagram - MODIFY Order Type
5.1 Product Name
5.1.1 Task Flow - Product Name
```

The supplied template's visual `Heading 1`, `Heading 2`, and `Heading 3` formatting is copied into numbering-free renderer styles. Level-4 content is rendered as bold Normal text because the supplied template does not contain a suitable Heading 4 style.

Business Rules and hierarchy bullets use the template's existing `List Paragraph` style rather than assuming the standard Word `List Bullet` style exists.

## Architecture

```
Excel (source)
     │
     ▼
extractors/          <- parses raw sheet structure into Python objects
     │                  (one module per table PATTERN, not per sheet)
     │                  all sheet lookups go through sheet_utils.find_sheet(),
     │                  which tolerates whitespace/case/typo differences
     │                  between the expected and actual sheet name
     ▼
models/schemas.py    <- pydantic contract every extractor must satisfy
     ▼
validation/checks.py <- business-rule sanity checks
     ▼
render/               <- reusable docx components + build_doc.py
                          which fills the Word master and injects
                          Excel/YAML-driven content
assets/process_guide_template.docx
                       <- approved Word document master; common format is
                          stored here instead of being recreated in Python
     ▼
output/ServiceDeliveryGuide.docx
```

## Document structure

When the supplied template is used, the final document is assembled as:

1. **Template front page** — preserved from the approved Word format and
   populated with product/document metadata.
2. **Template Document Control Sheet** — preserved layout; contact fields and
   first Revision History row are populated from the pipeline metadata.
3. **Template Contents page** — preserved layout with a live Word TOC field.
   Word is instructed to refresh fields when the document is opened.
4. **Section 1 Introduction** — the template's existing `Introduction` heading
   is used as the insertion anchor; introduction content is injected below it.
5. **Section 1.1 Product-related Business Rules** — Excel-driven.
6. **Section 2 Product Hierarchy** — Excel-driven, including hierarchy images.
7. **Section 2.1.x Product Attributes** — Excel-driven definitions and values.
8. **Section 3 Order Types with Descriptions** — YAML/manual source as before.
9. **Sections 4–N Workflow Diagrams** — dynamically generated from the order
   types found in the task-list sheet.
10. **Final Fault Handling Process** — stub until its source is identified.

The key design rule is: **Word controls presentation; Excel/YAML controls
variable content; Python controls the mapping/injection between them.**

## Task explanation pages — inclusion rule

A task gets its own explanation page (Task Owner / Description / Start
Event / End Event / KPI / OLA table + Business Rule / Work Instructions /
Required Information & Sources lines) UNLESS its Work Group is:
- blank
- exactly `SYSTEM` or `EXTERNAL` (case-insensitive)
- contains `auto` anywhere (case-insensitive) — catches `AUTO`,
  `AUTO (hidden)`, `AUTO <Hidden task>`, etc.

See `EXCLUDED_WORK_GROUPS_EXACT` and `_qualifies_for_explanation()` in
`render/build_doc.py` to adjust this rule.

## Current section status

| Section | Status | Source | Notes |
|---|---|---|---|
| Title Page, running header | generated | `--product` + CLI metadata flags | |
| Document Control Sheet | generated | CLI flags (`--contact-*`, etc.) | sign-off tables intentionally left blank for manual signing |
| 1 Introduction | generated* | `1. Product Description` | *unvalidated - no real sheet seen yet |
| 1.1 Business Rules | generated* | `3. Special Requirements` | *unvalidated |
| 2 Product Hierarchy + images | generated* | `2. Product Hierarchy` + images from `1. Product Description` | *unvalidated |
| 2.1.{i}.1 Attribute Definitions | generated* | same sheet, "Attribute" rows | *unvalidated |
| 2.1.{i}.2 Attribute List of Values | generated* | same sheet, "Value" rows following each attribute | *unvalidated - merged-cell table, see below |
| 3 Order Types summary table | generated | `5.1 MDM_Order Scenarios` | Excel is primary; YAML is fallback |
| 4–N Workflow Diagrams + task explanation pages | **generated & validated** | `6.3 OSS SO Task List` | tested against real data |
| Fault Handling Process | stub | not identified | |

**"Unvalidated" means:** built exactly to the mapping rules given, but
only tested against synthetic assumptions about layout — not a real
sheet, with 5.1 MDM_Order Scenarios and 3. Special Requirements now wired in. The sheet-name
resolver (`extractors/sheet_utils.py`) means a small typo or whitespace
difference in the real sheet name won't cause a false "not found" -
but the internal row/column layout assumptions (e.g. which column holds
"Attribute" vs "Product") are still unverified against a real file.

## Known data-quality issue (auto-detected, not fixed)

Row 80 of sheet `6.3 OSS SO Task List` has `"NO"` in the duration column
instead of a number. The pipeline doesn't crash on this — it sets
duration to `None` for that row and reports it as a warning every run.

## Sheet name matching

Every extractor resolves its expected sheet name through
`extractors/sheet_utils.find_sheet()`, which tries, in order:
1. Exact match
2. Case/whitespace-insensitive match
3. Match on the leading "N." number prefix only, if exactly one sheet
   shares that prefix (e.g. expecting "1. Product Description" still
   finds a sheet actually named "1. Product Descripton")

If none of these match, it raises a `ValueError` listing every real
sheet name in the workbook, so a genuine mismatch is always inspectable
rather than silently producing an empty section.

## Extending this pipeline

**To validate/fix the extractors for sheets 1, 2, 3 once you have real
data:** run the pipeline against a workbook containing them and read the
warnings - each extractor fails with a specific, actionable message
(missing sheet, missing column, no anchor rows found) rather than
silently producing wrong output. Adjust the relevant file in
`extractors/` to match the real layout, then add a regression test in
`tests/test_extractors.py` with known-good counts.

**To fill in Section 3's order-scenario table:** edit
`5.1 MDM_Order Scenarios` is now read automatically. `input/order_scenarios.yaml` remains available only as a fallback.

**To add a new order type in the future:** nothing needs to change.
Order types are discovered dynamically from merged title rows in sheet
6.3, and section numbering is computed, not hardcoded.

## Files

```
extractors/sheet_utils.py          - fuzzy sheet-name resolver, used by every other extractor
extractors/stacked_tables.py       - "title -> header -> data" stacked tables (6.3 pattern), product-parametrized
extractors/product_description.py  - sheet "1. Product Description" -> Section 1 (unvalidated)
extractors/special_requirements.py - sheet "3. Special Requirements" -> Section 1.1 (unvalidated)
extractors/product_hierarchy.py    - sheet "2. Product Hierarchy" -> Section 2 + 2.1.x (unvalidated)
extractors/hierarchy_images.py     - embedded images in sheet "1. Product Description" -> Section 2 (unvalidated)
models/schemas.py                  - pydantic contract: Task, OrderTypeSection, ProductComponent, Attribute, etc.
validation/checks.py               - business-rule checks run after extraction
render/components.py               - generic docx building blocks (heading, table, stub, TOC, horizontal rule)
render/template_utils.py           - {{PLACEHOLDER}} substitution across a template's body/header/footer
render/build_doc.py                - composes the full document: front page, header, Document Control Sheet, sections
pipeline_core.py                   - shared extract/render logic, used by both run.py and app.py
run.py                             - CLI entrypoint: extract -> validate -> render -> report
app.py                             - Streamlit web UI (upload Excel -> download .docx) for team/day-to-day use
input/order_scenarios.yaml         - optional fallback for Section 3's table
tests/test_extractors.py           - regression tests, incl. sheet-resolver and product-name-generalization tests
assets/placeholder_logo.png        - default logo used when --logo isn't provided
input/Book1.xlsx                   - source workbook (not bundled)
output/                            - generated documents land here
```

## Editable swimlane task-flow diagrams

For every order-type task flow, the renderer can generate an editable Word swimlane diagram from the same task table.

Rules:
- Tasks are sorted by `Task Seq.`.
- `Work Group` becomes the swimlane.
- If `Work Group` is blank, `IT System` is used as the swimlane.
- Tasks whose Work Group is exactly `External` or `System` (case-insensitive) are omitted from the diagram but remain in the source task table.
- The original four-column task table remains editable and is generated separately below the diagram.
- The diagram uses editable Word VML drawing objects rather than a flattened PNG/JPG. Microsoft Word desktop is the target renderer; third-party DOCX viewers may render legacy VML differently.
- Long flows are split into continuation panels so the task boxes remain readable.

Example structure:

```text
4. Workflow Diagram - CREATE Order Type
4.1 Product Name
4.1.1 Task Flow - Product Name
    Task Flow Diagram
    [editable swimlane diagram]
    Task Details
    [editable Task Seq. / Task Name / IT System / Work Group table]
```

To disable diagrams for a run, use:

```bash
python run.py ... --no-flowchart
```

## Independent editable flowchart shapes

The task-flow renderer creates each swimlane, task box, connector, start/end marker and continuation marker as an independent Word VML drawing object. The diagram is intentionally not placed inside a `v:group`, so shapes can be selected and moved independently in Microsoft Word.

The diagram uses a light-blue palette. Moving a task box does not automatically reroute or reattach its connector; connectors are independent editable objects and may be repositioned separately.


## Latest task-flow/document-detail updates

- Flowcharts are inserted directly under `4.1`, `5.1`, `6.1`, etc.
- The `4.1.1 Task Flow - <Product>` heading follows the diagram and introduces the task table.
- Task boxes are standard rectangles; Start/End are oval terminators; process symbols share a common size within each panel.
- Task sequence is used for ordering but is not printed in flowchart boxes.
- Sheet `6.4` is read for `Order Type`, `Task Name`, and `Work Order Description`; matching descriptions populate task detail pages.
- Sheet `5.1 MDM_Order Scenarios` is read for `Service Type`, `Order Sub Type`, `Condition`, and `Order Type` to populate Section 3.
- Sheet `3. Special Requirements` is rendered as an editable Word table when a source table is detected.
- `Task Owner` is populated from the task's Work Group.
- `Business Rule:`, `Work Instructions:`, and `Required Information & Sources:` are bold labels.
- Times New Roman is enforced across the final document.


### Latest extraction behavior

- Section 1.1 keeps Special Requirements as text and inserts only source tables belonging to the corresponding requirement.
- Repeated Attribute rows in Product Hierarchy are consolidated by attribute name and unique LOV values are merged.
