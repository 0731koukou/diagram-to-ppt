---
name: diagram-to-ppt
description: Use when users ask to convert or redraw PNG, JPG, JPEG, or screenshots of architecture diagrams, flowcharts, or organization charts into editable PowerPoint/PPTX files, especially when native shapes, editable text, movable connectors, annotations, or removal of full-page raster backgrounds are required.
---

# Diagram to PPT

## Overview

Reconstruct architecture diagrams, flowcharts, and organization charts as faithful native PowerPoint objects. Preserve both semantic structure and the reference's visual language; do not silently restyle a detailed diagram into a generic template.

This public skill requires a host LLM with multimodal image viewing, text reading, spatial reasoning, file execution, and rendered-preview inspection capabilities. Read [README.md](README.md) before use. The skill does not bundle OCR.

## Required Workflow

1. Inspect the reference image and classify it as architecture, flowchart, or organization chart.
2. Use the host LLM's visual capability to extract every visible label and build a text inventory containing the exact text, approximate bounding box, and confidence (`high` or `low`). Visually review the inventory before drawing. If the host cannot inspect the image directly, stop: the skill prerequisites are not met.
3. Read [references/reconstruction-workflow.md](references/reconstruction-workflow.md) and [references/visual-fidelity.md](references/visual-fidelity.md). Inventory visual style separately from semantic content.
4. Create a manifest that follows [references/manifest-schema.md](references/manifest-schema.md). Use source-image pixel coordinates, semantic native icons, independent title/body text layers, and explicit connector routes.
5. Render the manifest:

   ```powershell
   python scripts/render_manifest.py --manifest diagram.json --output diagram.pptx --summary build-summary.json
   ```

6. Run structural QA:

   ```powershell
   python scripts/inspect_pptx.py --pptx diagram.pptx --manifest diagram.json --output inspection.json
   ```

7. On Windows with PowerPoint installed, render and inspect actual text bounds:

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/powerpoint_render_qa.ps1 -PptxPath diagram.pptx -PreviewPath preview.png -JsonPath powerpoint-qa.json
   ```

8. Inspect the exported preview and iterate until [references/acceptance-rubric.md](references/acceptance-rubric.md) passes.

9. Generate a source/render contact sheet and advisory metrics:

   ```powershell
   python scripts/compare_previews.py --source source.png --candidate preview.png --output visual-qa.json --contact-sheet comparison.png
   ```

Do not claim general image-reconstruction readiness from synthetic manifests alone. For a real user request, run at least one user-provided or project-provided complex source image through the complete image-to-manifest-to-PPTX workflow.

## Non-Negotiable Rules

- Do not embed the full reference image in the final PPTX.
- Do not use raster crops as substitutes for nodes, labels, connectors, or ordinary diagram symbols.
- Keep all visible labels as PowerPoint text.
- Keep arrows and lines as independently selectable connectors.
- Match icon semantics. Use detailed native symbols or compose 5-12 primitives; do not replace a robot arm, brain, expert group, or report with an unrelated gear or rectangle.
- Preserve semantic hierarchy and connector direction before decorative details.
- Report when text in the source is unreadable; do not invent labels.
- Keep low-confidence text explicitly marked until visual review resolves it. Do not present uncertain visual transcription as confirmed source text.
- Treat mixed Chinese/Latin labels and compact captions as high-risk for PowerPoint wrapping. Fix actual geometry or font size until rendered overflow is zero; never silence the check.
- Do not accept a result only because structural QA passes. Reject visible downgrades in hierarchy, spacing, icon meaning, line routing, or object detail.

## Scope Boundary

Use this skill only for architecture diagrams, flowcharts, and organization charts. Do not use it for photographs, posters, general UI screenshots, scanned documents, or illustration-heavy artwork.

## Deliverables

Always deliver the editable `.pptx`, manifest, preview, source/render comparison, and QA JSON files. State native shape, text, connector, picture, missing-label, out-of-bounds, and text-overflow counts plus remaining visual simplifications.
