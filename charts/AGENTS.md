---
name: charts
description: C# examples for charts using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - charts

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **charts** category.
This folder contains standalone C# examples for charts operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **charts**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (31/31 files) ← category-specific
- `using Aspose.Words.Drawing.Charts;` (31/31 files)
- `using Aspose.Words.Drawing;` (30/31 files)
- `using System;` (24/31 files)
- `using System.Drawing;` (12/31 files)
- `using System.IO;` (4/31 files)
- `using Aspose.Words.Saving;` (1/31 files)
- `using System.Linq;` (1/31 files)
- `using Aspose.Words.Tables;` (1/31 files)

## Common Code Pattern

Most files follow this pattern:

```csharp
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
// ... operations ...
doc.Save("output.docx");
```

## Files in this folder

| File | Key APIs | Description |
|------|----------|-------------|
| [add-border-stroke-chart-legend-specified-thickness-dash...](./add-border-stroke-chart-legend-specified-thickness-dash-style-emphasis.cs) | `Aspose`, `Legend`, `Format` | Add border stroke chart legend specified thickness dash style emphasis |
| [add-multiple-chartdatapoint-objects-series-set-each-poi...](./add-multiple-chartdatapoint-objects-series-set-each-point-s-color-fill-property.cs) | `Color`, `Aspose`, `Document` | Add multiple chartdatapoint objects series set each point s color fill property |
| [add-scatter-chart-existing-paragraph-calling-documentbu...](./add-scatter-chart-existing-paragraph-calling-documentbuilder-insertchart-appropriate.cs) | `Aspose`, `Document`, `DocumentBuilder` | Add scatter chart existing paragraph calling documentbuilder insertchart appr... |
| [adjust-chart-legend-position-top-right-corner-set-its-b...](./adjust-chart-legend-position-top-right-corner-set-its-background-fill-light-gray.cs) | `Aspose`, `Document`, `DocumentBuilder` | Adjust chart legend position top right corner set its background fill light gray |
| [adjust-primary-y-axis-scaling-fixed-minimum-maximum-val...](./adjust-primary-y-axis-scaling-fixed-minimum-maximum-values-set-major-unit-intervals.cs) | `Aspose`, `Document`, `DocumentBuilder` | Adjust primary y axis scaling fixed minimum maximum values set major unit int... |
| [align-multi-line-chart-data-labels-center-enable-text-w...](./align-multi-line-chart-data-labels-center-enable-text-wrapping-better-readability.cs) | `Aspose`, `Document`, `DocumentBuilder` | Align multi line chart data labels center enable text wrapping better readabi... |
| [apply-predefined-chart-style-template-newly-inserted-ch...](./apply-predefined-chart-style-template-newly-inserted-chart-ensure-consistent-visual.cs) | `Title`, `Aspose`, `Document` | Apply predefined chart style template newly inserted chart ensure consistent... |
| [apply-three-dimensional-rotation-effect-column-chart-en...](./apply-three-dimensional-rotation-effect-column-chart-enhance-visual-perspective.cs) | `Aspose`, `DocumentBuilder`, `Words` | Apply three dimensional rotation effect column chart enhance visual perspective |
| [batch-process-folder-word-files-adding-predefined-bar-c...](./batch-process-folder-word-files-adding-predefined-bar-chart-each-document-s-first-page.cs) | `Aspose`, `Document`, `DocumentBuilder` | Batch process folder word files adding predefined bar chart each document s f... |
| [chartseriescollection-add-overload-accepting-name-value...](./chartseriescollection-add-overload-accepting-name-values-labeled-series-one-step.cs) | `Aspose`, `Document`, `DocumentBuilder` | Chartseriescollection add overload accepting name values labeled series one step |
| [clone-chart-shape-one-document-section-insert-duplicate...](./clone-chart-shape-one-document-section-insert-duplicate-another-paragraph.cs) | `Document`, `DocumentBuilder`, `Aspose` | Clone chart shape one document section insert duplicate another paragraph |
| [configure-chart-display-data-labels-only-points-exceedi...](./configure-chart-display-data-labels-only-points-exceeding-specified-threshold-value.cs) | `Aspose`, `Document`, `DocumentBuilder` | Configure chart display data labels only points exceeding specified threshold... |
| [customize-chart-s-data-label-font-specific-typeface-siz...](./customize-chart-s-data-label-font-specific-typeface-size-bold-styling-emphasis.cs) | `Aspose`, `Font`, `Document` | Customize chart s data label font specific typeface size bold styling emphasis |
| [define-default-chartdatalabel-options-apply-consistent-...](./define-default-chartdatalabel-options-apply-consistent-font-size-color-across-all.cs) | `DocumentBuilder`, `Aspose`, `Series` | Define default chartdatalabel options apply consistent font size color across... |
| [enable-automatic-resizing-chart-elements-when-document-...](./enable-automatic-resizing-chart-elements-when-document-page-size-changes-maintain.cs) | `Aspose`, `Series`, `Document` | Enable automatic resizing chart elements when document page size changes main... |
| [enable-data-label-leader-lines-pie-chart-customize-thei...](./enable-data-label-leader-lines-pie-chart-customize-their-length-better-placement.cs) | `Aspose`, `Document`, `DocumentBuilder` | Enable data label leader lines pie chart customize their length better placement |
| [existing-docx-file-locate-chart-shape-its-title-replace...](./existing-docx-file-locate-chart-shape-its-title-replace-its-data-source.cs) | `Aspose`, `Document`, `Words` | Existing docx file locate chart shape its title replace its data source |
| [implement-error-handling-catch-exceptions-when-insertin...](./implement-error-handling-catch-exceptions-when-inserting-chart-read-only-document.cs) | `Aspose`, `Document`, `DocumentBuilder` | Implement error handling catch exceptions when inserting chart read only docu... |
| [insert-chart-table-cell-ensure-it-scales-proportionally...](./insert-chart-table-cell-ensure-it-scales-proportionally-cell-dimensions.cs) | `Aspose`, `Words`, `Document` | Insert chart table cell ensure it scales proportionally cell dimensions |
| [insert-chart-two-dimensional-array-as-custom-data-sourc...](./insert-chart-two-dimensional-array-as-custom-data-source-mapping-series-categories.cs) | `Aspose`, `Document`, `DocumentBuilder` | Insert chart two dimensional array as custom data source mapping series categ... |
| [insert-column-chart-new-document-documentbuilder-insert...](./insert-column-chart-new-document-documentbuilder-insertchart-default-data.cs) | `Aspose`, `Document`, `DocumentBuilder` | Insert column chart new document documentbuilder insertchart default data |
| [new-chart-series-set-its-values-via-series-values-prope...](./new-chart-series-set-its-values-via-series-values-property-assign-custom-category.cs) | `Aspose`, `Document`, `DocumentBuilder` | New chart series set its values via series values property assign custom cate... |
| [programmatically-change-chart-s-plot-area-border-dashed...](./programmatically-change-chart-s-plot-area-border-dashed-line-specific-color-width.cs) | `DocumentBuilder`, `Aspose`, `Format` | Programmatically change chart s plot area border dashed line specific color w... |
| [programmatically-set-chart-s-background-fill-semi-trans...](./programmatically-set-chart-s-background-fill-semi-transparent-color-watermark-effect.cs) | `Aspose`, `Document`, `DocumentBuilder` | Programmatically set chart s background fill semi transparent color watermark... |
| [remove-specific-series-chart-chartseriescollection-remo...](./remove-specific-series-chart-chartseriescollection-removeat-correct-index.cs) | `Series`, `Aspose`, `Document` | Remove specific series chart chartseriescollection removeat correct index |
| [retrieve-existing-chart-series-modify-their-data-points...](./retrieve-existing-chart-series-modify-their-data-points-refresh-chart-display.cs) | `Aspose`, `YValues`, `Document` | Retrieve existing chart series modify their data points refresh chart display |
| [retrieve-shape-chart-object-inserted-chart-modify-its-t...](./retrieve-shape-chart-object-inserted-chart-modify-its-title-text-programmatically.cs) | `Aspose`, `Document`, `DocumentBuilder` | Retrieve shape chart object inserted chart modify its title text programmatic... |
| [set-secondary-y-axis-number-format-currency-two-decimal...](./set-secondary-y-axis-number-format-currency-two-decimal-places-financial-charts.cs) | `Aspose`, `Series`, `Document` | Set secondary y axis number format currency two decimal places financial charts |
| [update-chart-title-text-toggle-legend-visibility-based-...](./update-chart-title-text-toggle-legend-visibility-based-user-preferences.cs) | `Aspose`, `Document`, `Words` | Update chart title text toggle legend visibility based user preferences |
| [validate-that-all-chart-series-have-matching-category-c...](./validate-that-all-chart-series-have-matching-category-counts-prevent-data.cs) | `Series`, `Aspose`, `InvalidOperationException` | Validate that all chart series have matching category counts prevent data |
| ... | | *and 1 more files* |

## Category Statistics
- Total examples: 31

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for charts patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082255`
<!-- AUTOGENERATED:END -->
