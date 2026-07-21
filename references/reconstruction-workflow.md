# Reconstruction Workflow

## 1. Classify

- Architecture diagram: systems, zones, services, databases, interfaces, data flows.
- Flowchart: ordered steps, decisions, branches, loops, start/end states.
- Organization chart: parent-child hierarchy, reporting lines, departments, roles.

Reject photographs, posters, ordinary UI screenshots, scanned documents, and illustration-heavy images.

## 2. Inventory

Record:

- Canvas width, height, aspect ratio, background.
- Major containers and their bounds.
- Nodes in visual reading order.
- Exact visible text and confidence; mark unreadable text instead of guessing.
- Connector start, end, path, direction, line style, and arrowheads.
- Repeated styles: fill, border, corner type, typography, spacing.
- Symbols that can be simplified with native PowerPoint shapes.

Start with a lightweight text inventory before building shapes:

| Text | Approximate box `[x, y, w, h]` | Confidence | Review note |
|---|---|---|---|
| Exact visible label | Source-image pixels | `high` or `low` | Empty when confirmed |

Read text directly from the source with the host LLM's multimodal vision. Check each label against the image in context, especially Chinese characters, punctuation, digits, acronyms, and connector annotations. This skill does not bundle OCR; image reading is a required host capability. Keep unresolved text marked `low` or `[unreadable]` instead of guessing.

## 3. Calibrate Coordinates

Use the source image pixel dimensions as the manifest coordinate system. Keep nested elements in their parent region and reserve space for actual font rendering. Prefer explicit coordinates over automatic layout when matching a reference.

## 4. Build the Manifest

Create elements back-to-front: containers, nodes, symbols, text, then connectors and legends. Give every semantic object a stable `id` and `name`.

Use `shape.text` when text belongs inside a node. Use independent `text` elements for headings, captions, and labels that must move separately.

## 5. Render and Inspect

Run `render_manifest.py`, then `inspect_pptx.py`. Fix missing text, invalid element types, pictures, and out-of-bounds objects before visual review.

## 6. PowerPoint QA

Use `powerpoint_render_qa.ps1` when PowerPoint is available. It exports the real slide render and checks text bounds. Wait for PowerPoint layout to settle before export.

## 7. Visual Review

Compare the preview with the source in this order:

1. Major region placement and reading direction.
2. Node hierarchy and connector direction.
3. Text completeness and wrapping.
4. Alignment, spacing, colors, borders, and symbols.

Do not claim completion before structural and visual checks both pass.
