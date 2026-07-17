# SPOT Framework

**SPOT (Comment-Driven Worksheet Framework)** is an Excel VBA framework that transforms worksheet comments into structured application data.

Instead of treating Excel as a spreadsheet, SPOT treats a worksheet as a visual data editor.

Each **SPOT** represents a meaningful location on the worksheet. By combining comments, cell colors, and worksheet layouts, SPOT can describe maps, transitions, events, and other structured data.

The framework validates worksheet data before exporting it, allowing applications to work with reliable and reusable datasets.

## Features

* Comment-driven worksheet design
* Automatic validation before export
* Color-based area detection
* Transition generation
* Save / Load support
* Reusable structured data generation

SPOT is designed as a lightweight framework that allows Excel worksheets to become visual editors for games, simulations, and other data-driven applications.

## Why Comment-Driven?

Excel comments are more than annotations.

In SPOT, comments become structured metadata attached to worksheet cells.

A worksheet cell describes **where** something exists, while its comment describes **what** it represents.

For example:

```text
Cell
    Position
    Size
    Color

Comment
    SLOT:1
    DATA:1001
    DIR:DOWN
```

This separation allows worksheet layouts and application data to remain independent.

SPOT reads worksheet comments, validates their contents, detects related worksheet areas, and transforms them into reusable structured data.

Because the data is stored as comments instead of visible worksheet values, the worksheet remains clean and easy to edit while still containing rich application metadata.

This approach makes the same worksheet suitable for maps, transitions, simulations, events, user interfaces, and many other structured data models.

## Quick Start

The workbook contains two sample worksheets:

* **Sample1** – Paired Area (TYPE_A)
* **Sample2** – Transition Data (TYPE_B)

Use **Sheet1** as the working area.

### TYPE_A

1. Copy **Sample1** to **Sheet1**.
2. Run **CopyPairAreas**.
3. Add comments, SLOT numbers, and DATA IDs.
4. Run **Validate_TYPE_A**.
5. Run **Save**.

### TYPE_B

1. Copy **Sample2** to **Sheet1**.
2. Edit transition comments.
3. Run **Validate_TYPE_B**.

The sample worksheets are templates.

Always perform editing and validation in **Sheet1** to keep the original samples unchanged.

## Architecture

SPOT follows a simple processing pipeline.

```text
Worksheet
    ↓
Validate
    ↓
Detect
    ↓
Generate Structured Data
    ↓
Application / Engine
```

### Worksheet

A worksheet acts as a visual editor.

Each **SPOT** consists of worksheet cells and their associated comments.

### Validate

Validation checks ensure that worksheet data is complete and consistent before any data is generated.

Examples include:

* Missing SLOT numbers
* Invalid comment formats
* Rectangle validation
* Transition validation

### Detect

The detection stage analyzes worksheet information such as:

* Colored areas
* Cell positions
* Comments
* SLOT relationships

### Generate Structured Data

Validated worksheet data is transformed into reusable structured data that can be consumed by applications, game engines, simulation systems, or other external tools.

SPOT separates editing from execution.

Excel is used for creating and validating data, while applications are responsible for rendering, simulation, and runtime behavior.

## Examples

SPOT is designed to generate reusable structured data.

Possible applications include:

- Spot the Difference games
- Tile map editors
- Map transition editors
- Event editors
- User interface layout editors
- Simulation editors
- Visual data editors

