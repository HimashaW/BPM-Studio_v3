from __future__ import annotations

import json
import os
from typing import Any

from google import genai


def _source_context(result) -> dict[str, Any]:
    """Build a compact, source-grounded context from the Python extraction result."""
    hierarchy = []
    for node in result.hierarchy:
        attrs = []
        for attr in node.attributes:
            attrs.append({
                "name": attr.entity_name,
                "mandatory": attr.mandatory,
                "comments": attr.comments,
                "values": attr.values,
            })
        hierarchy.append({"kind": node.kind, "name": node.name, "attributes": attrs})

    rules = [b.text for b in result.business_rule_blocks if b.text]
    if not rules:
        rules = result.business_rules

    scenarios = [
        {
            "business_scenario": r.business_scenario,
            "crm_order_sub_type": r.crm_order_sub_type,
            "crm_order_condition": r.crm_order_condition,
            "oss_service_order": r.oss_service_order,
        }
        for r in result.order_scenarios
    ]

    return {
        "workbook_name": result.workbook_name,
        "product_name": result.product_name,
        "product_description_source": result.introduction_text,
        "order_types": [o.order_type for o in result.order_types],
        "product_hierarchy": hierarchy,
        "business_rules": rules,
        "order_scenarios": scenarios,
    }


def generate_introduction(result, api_key: str | None = None, model: str = "gemini-3.5-flash") -> list[str]:
    """Generate a professional introduction using only facts extracted from the workbook."""
    api_key = api_key or os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY is not set. Add it to the environment before using --ai-enhance.")

    client = genai.Client(api_key=api_key)
    context = _source_context(result)

    prompt = f"""
You are an expert telecommunications Service Delivery Guide writer.

The Python pipeline has extracted the following information directly from the source Excel workbook.
The workbook is the source of truth.

SOURCE DATA:
{json.dumps(context, ensure_ascii=False, indent=2)}

Create a professional Introduction for a Service Delivery Guide.

Return ONLY valid JSON in exactly this structure:
{{
  "paragraphs": ["paragraph 1", "paragraph 2", "paragraph 3"],
  "order_types": ["CREATE"]
}}

Rules:
- Use only facts supported by SOURCE DATA.
- Do not invent capabilities, technologies, SLAs, processes, benefits, or customer claims.
- Do not copy the Excel wording verbatim; rewrite it professionally.
- Explain the service/product at a high level.
- Explain the purpose and scope of the guide.
- Mention relevant order types when useful.
- order_types must contain only order types present in SOURCE DATA.
- Do not turn detailed business rules or attribute LOVs into the introduction.
- Do not create tables.
- Do not mention that AI was used.
- Keep the introduction concise: 3 to 5 paragraphs.
"""

    response = client.models.generate_content(
        model=model,
        contents=prompt,
        config={"response_mime_type": "application/json"},
    )

    data = json.loads(response.text)
    paragraphs = data.get("paragraphs")
    if not isinstance(paragraphs, list) or not all(isinstance(p, str) and p.strip() for p in paragraphs):
        raise ValueError("Gemini returned an invalid Introduction structure.")

    source_orders = {o.order_type for o in result.order_types}
    returned_orders = data.get("order_types", [])
    if not isinstance(returned_orders, list) or not all(isinstance(x, str) for x in returned_orders):
        raise ValueError("Gemini returned an invalid order_types structure.")
    unknown = set(returned_orders) - source_orders
    if unknown:
        raise ValueError(f"Gemini returned unsupported order type(s): {sorted(unknown)}")

    return [p.strip() for p in paragraphs]
