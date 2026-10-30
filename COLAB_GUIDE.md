# Google Colab — Generate a Service Delivery Guide

## 1. Upload the pipeline ZIP

```python
from google.colab import files
uploaded = files.upload()
```

Choose the latest pipeline ZIP.

## 2. Extract it

```python
import zipfile

zip_name = next(iter(uploaded.keys()))
with zipfile.ZipFile(zip_name, "r") as z:
    z.extractall("/content")

PIPELINE_DIR = "/content/bpm-userguide-pipeline"
print("Pipeline:", PIPELINE_DIR)
```

## 3. Install dependencies

```python
%cd /content/bpm-userguide-pipeline
!pip install -q -r requirements.txt
```

## 4. Upload the Excel workbook

Do not hard-code the filename. Colab may add `(1)`, `(2)`, etc. if the same file is uploaded more than once.

```python
from google.colab import files
import os

uploaded_excel = files.upload()
excel_name = next(iter(uploaded_excel.keys()))

input_dir = "/content/bpm-userguide-pipeline/input"
os.makedirs(input_dir, exist_ok=True)

excel_path = os.path.join(input_dir, excel_name)
with open(excel_path, "wb") as f:
    f.write(uploaded_excel[excel_name])

print("Excel:", excel_path)
```

## 5. Verify the workbook

```python
import openpyxl

wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
print("Sheets:")
for i, name in enumerate(wb.sheetnames, 1):
    print(f"{i}. {name!r}")
```

The pipeline expects the task list in the `6.3` sheet and task descriptions in a `6.4` sheet containing these columns:

- Order Type
- Task Name
- Work Order Description

Minor sheet-name whitespace differences are handled.

## 6. Generate the DOCX

```python
%cd /content/bpm-userguide-pipeline

!python run.py \
    --input "$excel_path" \
    --output "/content/generated_Service_Delivery_Guide.docx" \
    --product "MS-Edge Defender" \
    --doc-no "BPMS-MS-EDGE-001" \
    --issue-no "001" \
    --issue-date "11.08.2026" \
    --revision-no "0" \
    --revision-date "11.08.2026" \
    --issued-by "Business Process Management Section"
```

The bundled Word file at `assets/process_guide_template.docx` is used automatically.

## 7. Download

```python
from google.colab import files
files.download("/content/generated_Service_Delivery_Guide.docx")
```

## Flowchart rules

For each order type:

```text
4. Workflow Diagram - CREATE Order Type
4.1 MS-Edge Defender
    [editable swimlane flowchart]
4.1.1 Task Flow - MS-Edge Defender
    [editable task table]
```

The next order type becomes 5 / 5.1 / 5.1.1, etc.

The flowchart:

1. Sorts tasks by `Task Seq.`.
2. Uses `Work Group` as the swimlane.
3. If Work Group is blank, uses `IT System` as the swimlane.
4. Omits Work Group `External` and `System` from the diagram only.
5. Does not print the sequence number inside a task shape.
6. Uses light-blue standard process rectangles and oval Start/End shapes.
7. Uses the same process-symbol width and height within a panel.
8. Keeps task boxes, lanes, connectors and Start/End as independent editable Word objects.

Microsoft Word desktop is the intended renderer. Moving a task shape does not automatically reroute its independent connector.

## Task detail pages

For each qualifying task, the generated detail table contains:

- Task Owner — populated from Work Group
- Description — populated from the matching 6.4 `Work Order Description`
- Start Event — blank
- End Event — blank
- KPI — blank
- OLA — blank

The labels `Business Rule:`, `Work Instructions:` and `Required Information & Sources:` are bold and their content areas remain blank for manual completion.

## Font

The final document is forced to Times New Roman throughout, including the supplied template's existing text, headers/footers, generated tables and generated flowchart text.

## Disable flowcharts

Add this option if needed:

```bash
--no-flowchart
```


### Latest extraction behavior

- Section 1.1 keeps Special Requirements as text and inserts only source tables belonging to the corresponding requirement.
- Repeated Attribute rows in Product Hierarchy are consolidated by attribute name and unique LOV values are merged.

## 8. Enable Gemini AI enhancement (real Excel data)

The AI enhancement uses the same `ExtractionResult` produced from the uploaded Excel workbook. No sample/manual `product_context` is required.

If your Gemini key is stored in Colab Secrets, load it with:

```python
from google.colab import userdata
import os

os.environ["GEMINI_API_KEY"] = userdata.get("GEMINI_API_KEY")
```

Or, if you already set the environment variable earlier, verify it without printing the key:

```python
import os
print("Gemini key available:", bool(os.environ.get("GEMINI_API_KEY")))
```

Install the updated dependencies:

```python
%cd /content/bpm-userguide-pipeline
!pip install -q -r requirements.txt
```

Then run the real pipeline with Gemini enabled:

```python
!python run.py \
    --input "$excel_path" \
    --output "/content/generated_AI_Service_Delivery_Guide.docx" \
    --product "MS-Edge Defender" \
    --doc-no "BPMS-MS-EDGE-001" \
    --issue-no "001" \
    --issue-date "05.10.2026" \
    --revision-no "0" \
    --revision-date "05.10.2026" \
    --issued-by "Business Process Management Section" \
    --ai-enhance \
    --ai-model "gemini-3.8-flash"
```

With `--ai-enhance`, the pipeline follows:

```text
Uploaded Excel
    -> existing Python extraction
    -> validated ExtractionResult
    -> Gemini (Introduction enhancement)
    -> Python validation
    -> existing Word template renderer
    -> AI-enhanced Service Delivery Guide
```

If Gemini fails, the pipeline does not stop. It records a warning and falls back to the original Excel-derived Introduction.

### Important

At this stage Gemini only enhances the Introduction. The next AI phase will use the **same real extracted tasks from 6.3/6.4** to produce structured workflow interpretation for the editable Word flowcharts. The deterministic Python rules remain the source of truth.
