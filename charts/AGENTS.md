---
name: charts
description: Verified C# examples for charts scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Charts

## Purpose

This folder is a **live, curated example set** for charts scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free chart insertion, configuration, validation, and export workflows using direct Aspose.Words APIs.

## Non-negotiable conventions

- Use `Shape` + `Chart` APIs directly and validate `shape.HasChart` before modification.
- Insert charts with `DocumentBuilder.InsertChart` using the required `ChartType`.
- Bootstrap local sample DOCX inputs for existing-file, folder, and stream workflows.
- Prefer simple, verifiable chart workflows over speculative formatting shortcuts.
- Guard maybe-null values to avoid nullable-reference warnings such as `CS8600`, `CS8602`, and `CS8604`.

## Recommended workflow selection

- **Native chart API workflow**: 28 examples
- **Existing DOCX / export workflow**: 5 examples
- **Validation workflow**: 3 examples
- **Stream / batch / input-bootstrap workflow**: 3 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. Chart operations must target actual chart shapes and preserve valid chart data structure.
3. Exported outputs (DOCX/PDF/HTML/etc.) must actually be written by the example.
4. Validation scenarios must inspect actual chart series and point collections.
5. Examples that depend on files, folders, streams, or existing-docx style inputs should bootstrap those inputs locally during the example run.

## File-to-task reference

- `load-an-existing-docx-file-locate-a-chart-shape-by-its-title-and-replace-its-data-source.cs`
  - Task: Load an existing DOCX file, locate a chart shape by its title, and replace its data source.
  - Workflow: existing-docx-and-export
  - Outputs: docx
  - Selected engine: verified
- `insert-a-column-chart-into-a-new-document-using-documentbuilder-insertchart-with-default-d.cs`
  - Task: Insert a column chart into a new document using DocumentBuilder.InsertChart with default data.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `add-a-scatter-chart-to-an-existing-paragraph-by-calling-documentbuilder-insertchart-with-t.cs`
  - Task: Add a scatter chart to an existing paragraph by calling DocumentBuilder.InsertChart with the appropriate overload.
  - Workflow: existing-docx-and-export
  - Outputs: docx
  - Selected engine: verified
- `insert-a-chart-using-a-two-dimensional-array-as-a-custom-data-source-mapping-series-and-ca.cs`
  - Task: Insert a chart using a two‑dimensional array as a custom data source, mapping series and categories.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `insert-a-chart-into-a-table-cell-and-ensure-it-scales-proportionally-with-the-cell-dimensi.cs`
  - Task: Insert a chart into a table cell and ensure it scales proportionally with the cell dimensions.
  - Workflow: validation
  - Outputs: docx
  - Selected engine: verified
- `clone-a-chart-shape-from-one-document-section-and-insert-the-duplicate-into-another-paragr.cs`
  - Task: Clone a chart shape from one document section and insert the duplicate into another paragraph.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `retrieve-the-shape-chart-object-from-an-inserted-chart-and-modify-its-title-text-programma.cs`
  - Task: Retrieve the Shape.Chart object from an inserted chart and modify its title text programmatically.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `create-a-new-chart-series-set-its-values-via-the-series-values-property-and-assign-custom.cs`
  - Task: Create a new chart series, set its values via the series.Values property, and assign custom category labels.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `use-chartseriescollection-add-overload-accepting-a-name-and-values-to-create-a-labeled-ser.cs`
  - Task: Use ChartSeriesCollection.Add overload accepting a name and values to create a labeled series in one step.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `add-multiple-chartdatapoint-objects-to-a-series-and-set-each-point-s-color-using-the-fill.cs`
  - Task: Add multiple ChartDataPoint objects to a series and set each point's color using the Fill property.
  - Workflow: stream-batch-io
  - Outputs: docx
  - Selected engine: verified
- `change-the-chart-type-from-column-to-line-after-populating-data-to-demonstrate-dynamic-tra.cs`
  - Task: Change the chart type from column to line after populating data to demonstrate dynamic transformation.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `add-a-trendline-to-a-scatter-chart-series-and-configure-its-type-color-and-display-equatio.cs`
  - Task: Add a trendline to a scatter chart series and configure its type, color, and display equation.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `configure-chartdatalabel-number-format-to-display-percentages-with-one-decimal-place-for-a.cs`
  - Task: Configure ChartDataLabel number format to display percentages with one decimal place for all series data points.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `align-multi-line-chart-data-labels-to-the-center-and-enable-text-wrapping-for-better-reada.cs`
  - Task: Align multi‑line chart data labels to the center and enable text wrapping for better readability.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `define-default-chartdatalabel-options-to-apply-consistent-font-size-and-color-across-all-c.cs`
  - Task: Define default ChartDataLabel options to apply consistent font size and color across all chart series.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `customize-the-chart-s-data-label-font-to-use-a-specific-typeface-size-and-bold-styling-for.cs`
  - Task: Customize the chart's data label font to use a specific typeface, size, and bold styling for emphasis.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `enable-data-label-leader-lines-for-a-pie-chart-and-customize-their-length-for-better-place.cs`
  - Task: Enable data label leader lines for a pie chart and customize their length for better placement.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `configure-the-chart-to-display-data-labels-only-for-points-exceeding-a-specified-threshold.cs`
  - Task: Configure the chart to display data labels only for points exceeding a specified threshold value.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `set-major-gridlines-visibility-on-the-primary-x-axis-and-customize-their-line-color-and-th.cs`
  - Task: Set major gridlines visibility on the primary X‑axis and customize their line color and thickness.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `adjust-the-primary-y-axis-scaling-to-fixed-minimum-and-maximum-values-and-set-major-unit-i.cs`
  - Task: Adjust the primary Y‑axis scaling to fixed minimum and maximum values and set major unit intervals.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `set-display-units-for-the-secondary-x-axis-to-thousands-and-format-axis-labels-with-a-cust.cs`
  - Task: Set display units for the secondary X‑axis to thousands and format axis labels with a custom number format.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `set-the-secondary-y-axis-number-format-to-currency-with-two-decimal-places-for-financial-c.cs`
  - Task: Set the secondary Y‑axis number format to currency with two decimal places for financial charts.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `apply-a-solid-fill-color-to-the-chart-plot-area-and-add-a-gradient-overlay-for-visual-dept.cs`
  - Task: Apply a solid fill color to the chart plot area and add a gradient overlay for visual depth.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `programmatically-set-the-chart-s-background-fill-to-a-semi-transparent-color-to-create-a-w.cs`
  - Task: Programmatically set the chart's background fill to a semi‑transparent color to create a watermark effect.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `add-a-border-stroke-to-the-chart-legend-with-specified-thickness-and-dash-style-for-emphas.cs`
  - Task: Add a border stroke to the chart legend with specified thickness and dash style for emphasis.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `adjust-the-chart-legend-position-to-the-top-right-corner-and-set-its-background-fill-to-li.cs`
  - Task: Adjust the chart legend position to the top right corner and set its background fill to light gray.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `update-the-chart-title-text-and-toggle-legend-visibility-based-on-user-preferences.cs`
  - Task: Update the chart title text and toggle legend visibility based on user preferences.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `apply-a-predefined-chart-style-template-to-a-newly-inserted-chart-to-ensure-consistent-vis.cs`
  - Task: Apply a predefined chart style template to a newly inserted chart to ensure consistent visual branding.
  - Workflow: validation
  - Outputs: docx
  - Selected engine: verified
- `apply-a-three-dimensional-rotation-effect-to-a-column-chart-to-enhance-visual-perspective.cs`
  - Task: Apply a three‑dimensional rotation effect to a column chart to enhance visual perspective.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `programmatically-change-the-chart-s-plot-area-border-to-a-dashed-line-with-specific-color.cs`
  - Task: Programmatically change the chart's plot area border to a dashed line with specific color and width.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `programmatically-hide-the-chart-s-plot-area-border-while-keeping-axis-lines-visible-for-a.cs`
  - Task: Programmatically hide the chart's plot area border while keeping axis lines visible for a clean appearance.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `enable-automatic-resizing-of-chart-elements-when-the-document-page-size-changes-to-maintai.cs`
  - Task: Enable automatic resizing of chart elements when the document page size changes to maintain layout integrity.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `retrieve-existing-chart-series-modify-their-data-points-and-refresh-the-chart-display.cs`
  - Task: Retrieve existing chart series, modify their data points, and refresh the chart display.
  - Workflow: existing-docx-and-export
  - Outputs: docx
  - Selected engine: verified
- `remove-a-specific-series-from-a-chart-using-chartseriescollection-removeat-with-the-correc.cs`
  - Task: Remove a specific series from a chart using ChartSeriesCollection.RemoveAt with the correct index.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: verified
- `validate-that-a-chart-contains-the-expected-number-of-series-and-data-points-before-saving.cs`
  - Task: Validate that a chart contains the expected number of series and data points before saving.
  - Workflow: validation
  - Outputs: docx
  - Selected engine: verified
- `validate-that-all-chart-series-have-matching-category-counts-to-prevent-data-misalignment.cs`
  - Task: Validate that all chart series have matching category counts to prevent data misalignment errors during rendering.
  - Workflow: existing-docx-and-export
  - Outputs: docx
  - Selected engine: verified
- `export-a-word-document-containing-multiple-charts-to-pdf-while-preserving-chart-formatting.cs`
  - Task: Export a Word document containing multiple charts to PDF while preserving chart formatting and data labels.
  - Workflow: stream-batch-io
  - Outputs: pdf, docx
  - Selected engine: verified
- `implement-error-handling-to-catch-exceptions-when-inserting-a-chart-into-a-read-only-docum.cs`
  - Task: Implement error handling to catch exceptions when inserting a chart into a read‑only document stream.
  - Workflow: existing-docx-and-export
  - Outputs: docx
  - Selected engine: verified
- `batch-process-a-folder-of-word-files-adding-a-predefined-bar-chart-to-each-document-s-firs.cs`
  - Task: Batch process a folder of Word files, adding a predefined bar chart to each document's first page.
  - Workflow: stream-batch-io
  - Outputs: docx
  - Selected engine: verified

## Common failure patterns and preferred agent fixes

- **Invalid chart shape assumption**
  - Symptom: Code treats a generic Shape as a chart without checking HasChart.
  - Preferred fix: Validate shape.HasChart before accessing the Chart property.

- **Series/data misalignment**
  - Symptom: Category count and series values do not match expected chart structure.
  - Preferred fix: Clear default data if needed and add aligned categories and series values deterministically.

- **Invalid table insertion state**
  - Symptom: Chart insertion in a table cell fails because the builder is left in an unbalanced table state.
  - Preferred fix: Balance StartTable, InsertCell, EndRow, and EndTable before saving.

- **Missing local bootstrap inputs**
  - Symptom: Existing DOCX, folder, or stream scenarios assume inputs already exist.
  - Preferred fix: Create local sample input documents or folders inside the example first.

- **Nullable warnings**
  - Symptom: CS8600, CS8602, or CS8604 around maybe-null chart shapes or located nodes.
  - Preferred fix: Use nullable locals and guard maybe-null values before dereference.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\charts\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer direct Aspose.Words APIs over speculative shortcuts.
