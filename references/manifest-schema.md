# Manifest Schema

## Top Level

```json
{
  "deck": {
    "name": "Editable diagram",
    "canvas_width": 1600,
    "canvas_height": 900,
    "slide_width_in": 13.333,
    "slide_height_in": 7.5
  },
  "slides": [
    {
      "name": "Diagram",
      "background": "#FFFFFF",
      "elements": []
    }
  ]
}
```

Use hexadecimal colors. Coordinates are source-image pixels.

## Shape

```json
{
  "type": "shape",
  "id": "service-a",
  "name": "Service A",
  "shape": "roundRect",
  "x": 120, "y": 180, "w": 260, "h": 90,
  "fill": "#F7FBFF",
  "stroke": "#1677FF",
  "stroke_width_pt": 1.2,
  "dash": false,
  "text": "Service A",
  "font_family": "Microsoft YaHei",
  "font_size_pt": 16,
  "bold": true,
  "color": "#10213A",
  "line_spacing": 1.0,
  "align": "center",
  "valign": "middle"
}
```

Supported shapes: `rect`, `roundRect`, `ellipse`, `diamond`, `hexagon`, `chevron`, `parallelogram`, `cylinder`, `cube`, `cloud`, `terminator`.

## Text

```json
{
  "type": "text",
  "id": "title",
  "x": 60, "y": 30, "w": 800, "h": 60,
  "text": "System Architecture",
  "font_family": "Microsoft YaHei",
  "font_size_pt": 26,
  "bold": true,
  "color": "#10213A",
  "align": "left",
  "valign": "middle"
}
```

## Connector

```json
{
  "type": "connector",
  "id": "service-a-to-db",
  "points": [[380, 225], [520, 225], [520, 410]],
  "stroke": "#1677FF",
  "stroke_width_pt": 1.5,
  "dash": false,
  "arrow_start": false,
  "arrow_end": true
}
```

Each point pair becomes an independently editable PowerPoint line segment. Arrowheads are applied to the first or last segment.

## Symbol

```json
{
  "type": "symbol",
  "id": "db-icon",
  "symbol": "database",
  "x": 560, "y": 365, "w": 48, "h": 54,
  "color": "#16A085"
}
```

Supported symbols: `arm`, `monitor`, `clipboard`, `clipboard-check`, `book`, `person`, `people`, `brain`, `cube`, `report`, `database`, `warehouse`, `tools`, `bars`, `refresh`, `document`, `server`, and `gear`.

Detailed semantic symbols are composed entirely from native PowerPoint shapes and line segments. Size them with the `x`, `y`, `w`, and `h` bounding box.

Shapes also accept `fill_transparency` from 0 to 100. Text and text-bearing shapes accept `line_spacing`.

## Prohibited Element

The renderer rejects `image`, `picture`, and raster component elements. Use another workflow when raster content is essential.
