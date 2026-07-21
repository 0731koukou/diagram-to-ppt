# Acceptance Rubric

The result passes only when all applicable checks succeed.

## Structure

- The PPTX opens successfully.
- The slide aspect ratio matches the source image.
- All major containers, nodes, labels, and connectors exist.
- Connector directions and branch relationships are correct.

## Native Editability

- Picture shapes: 0.
- Files under `ppt/media/`: 0.
- Labels are editable PowerPoint text.
- Nodes are editable PowerPoint shapes.
- Connectors are independently selectable line segments.

## Layout

- Missing required text: 0.
- Out-of-bounds shapes: 0.
- PowerPoint text overflow: 0.
- No visible text collision, clipping, or incoherent overlap.

## Visual Review

- Major region placement matches at first glance.
- Reading order and hierarchy match the source.
- Colors, line styles, spacing, and typography are consistent.
- Simplified symbols retain their semantic meaning.
- Intentional title/body line breaks match the reference.
- Major icons have the correct meaning and comparable native detail.
- A full-size source/render/difference contact sheet has been inspected.
- Passing structural checks does not override an obvious first-glance visual downgrade.
- A real source image, not only a hand-authored synthetic fixture, has completed the full workflow before generalized readiness is claimed.

## Deliverables

- Editable `.pptx`.
- Source manifest `.json`.
- Structural QA `.json`.
- PowerPoint QA `.json` when PowerPoint is available.
- Rendered preview `.png`.
