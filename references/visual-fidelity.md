# Visual Fidelity

## Default Quality Level

Use **faithful native redraw** unless the user explicitly accepts simplification. Preserve the reference's visual language with editable PowerPoint primitives. A structurally correct generic template is not a faithful redraw.

Do not use this skill for pixel-identical raster overlays. If exact appearance requires the source image as a background, report that it conflicts with full native editability.

## Inventory Before Rendering

Record these separately from semantic content:

| Surface | Capture |
| --- | --- |
| Canvas | aspect ratio, margins, occupied bounds, whitespace |
| Palette | background, region, role, connector, muted text colors |
| Typography | title/body hierarchy, line breaks, Chinese/Latin sizing |
| Geometry | corner style, border weight, dashed pattern, concentric rings |
| Icons | semantic subject, primitive count, stroke style, fill behavior |
| Connectors | topology, routing lanes, bidirectionality, feedback style |
| Layering | grid, regions, connectors, cards, icons, text, masks, legend |

## Native Icon Rule

Match meaning before convenience. Do not substitute a gear for a robot arm, a brain, a document, or an expert group merely because the color matches.

Use a supported semantic symbol when available. Otherwise compose an icon from native shapes and line segments. Prefer 5–12 purposeful primitives over one unrelated generic symbol.

Supported detailed symbols include `arm`, `monitor`, `clipboard`, `clipboard-check`, `book`, `people`, `brain`, `cube`, `report`, `database`, `warehouse`, `tools`, `bars`, and `refresh`.

## Text Rule

- Keep title and each body line in independent text boxes when their sizes, positions, or wrapping risks differ.
- Preserve intentional source line breaks.
- Reserve width using the actual PowerPoint font renderer; mixed Chinese/Latin labels need extra room.
- Do not reduce body text until it becomes visually subordinate to the reference merely to clear overflow.

## Layering Rule

Build in this order unless the reference proves otherwise:

1. background and optional subtle grid;
2. major regions;
3. connector routing lanes;
4. cards and core shapes;
5. semantic icons;
6. independent text layers;
7. explicit masks for interrupted dashed borders;
8. notes and legend.

## Visual QA

Run `compare_previews.py` after PowerPoint export. Inspect the contact sheet at full slide size. Metrics are advisory because white backgrounds inflate pixel similarity.

Reject the result when any of these are true:

- a major icon has the wrong meaning;
- title/body hierarchy or line breaks visibly differ;
- connector routes cross labels or lose feedback direction;
- whitespace, card proportions, or region balance are materially different;
- a detailed reference is reduced to a visibly generic template;
- the result passes overflow checks but still looks worse at first glance.

When an editable benchmark exists, compare object granularity as a diagnostic. A large object-count drop usually signals lost icon detail, text layers, connector segments, or masks; it is not a target by itself.
