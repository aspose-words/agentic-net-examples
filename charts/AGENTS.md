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

- `using Aspose.Words;` (30/31 files) ← category-specific
- `using Aspose.Words.Drawing.Charts;` (30/31 files)
- `using Aspose.Words.Drawing;` (29/31 files)
- `using System;` (24/31 files)
- `using System.Drawing;` (11/31 files)
- `using System.IO;` (6/31 files)
- `using Aspose.Words.Tables;` (2/31 files)
- `using Aspose.Words.Saving;` (1/31 files)
- `using System.Linq;` (1/31 files)

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
| [add-border-stroke-chart-legend-specified-thickness-dash...](./add-border-stroke-chart-legend-specified-thickness-dash-style-emphasis.cs) | `Legend`, `Format`, `Stroke` | Add border stroke chart legend specified thickness dash style emphasis |
| [add-multiple-chartdatapoint-objects-series-set-each-poi...](./add-multiple-chartdatapoint-objects-series-set-each-point-s-color-fill-property.cs) | `Color`, `Document`, `DocumentBuilder` | Add multiple chartdatapoint objects series set each point s color fill property |
| [add-scatter-chart-existing-paragraph-calling-documentbu...](./add-scatter-chart-existing-paragraph-calling-documentbuilder-insertchart-appropriate.cs) | `Document`, `DocumentBuilder`, `Series` | Add scatter chart existing paragraph calling documentbuilder insertchart appr... |
| [adjust-chart-legend-position-top-right-corner-set-its-b...](./adjust-chart-legend-position-top-right-corner-set-its-background-fill-light-gray.cs) | `Document`, `DocumentBuilder`, `Drawing` | Adjust chart legend position top right corner set its background fill light gray |
| [adjust-primary-y-axis-scaling-fixed-minimum-maximum-val...](./adjust-primary-y-axis-scaling-fixed-minimum-maximum-values-set-major-unit-intervals.cs) | `Document`, `DocumentBuilder`, `AxisBound` | Adjust primary y axis scaling fixed minimum maximum values set major unit int... |
| [align-multi-line-chart-data-labels-center-enable-text-w...](./align-multi-line-chart-data-labels-center-enable-text-wrapping-better-readability.cs) | `Document`, `DocumentBuilder`, `Series` | Align multi line chart data labels center enable text wrapping better readabi... |
| [apply-predefined-chart-style-template-newly-inserted-ch...](./apply-predefined-chart-style-template-newly-inserted-chart-ensure-consistent-visual.cs) | `Title`, `Document`, `DocumentBuilder` | Apply predefined chart style template newly inserted chart ensure consistent... |
| [apply-three-dimensional-rotation-effect-column-chart-en...](./apply-three-dimensional-rotation-effect-column-chart-enhance-visual-perspective.cs) | `DocumentBuilder`, `Document`, `Series` | Apply three dimensional rotation effect column chart enhance visual perspective |
| [batch-process-folder-word-files-adding-predefined-bar-c...](./batch-process-folder-word-files-adding-predefined-bar-chart-each-document-s-first-page.cs) | `Document`, `DocumentBuilder`, `Series` | Batch process folder word files adding predefined bar chart each document s f... |
| [chartseriescollection-add-overload-accepting-name-value...](./chartseriescollection-add-overload-accepting-name-values-labeled-series-one-step.cs) | `Document`, `DocumentBuilder`, `Series` | Chartseriescollection add overload accepting name values labeled series one step |
| [clone-chart-shape-one-document-section-insert-duplicate...](./clone-chart-shape-one-document-section-insert-duplicate-another-paragraph.cs) | `Document`, `DocumentBuilder`, `ConvertUtil` | Clone chart shape one document section insert duplicate another paragraph |
| [configure-chart-display-data-labels-only-points-exceedi...](./configure-chart-display-data-labels-only-points-exceeding-specified-threshold-value.cs) | `Document`, `DocumentBuilder`, `Series` | Configure chart display data labels only points exceeding specified threshold... |
| [customize-chart-s-data-label-font-specific-typeface-siz...](./customize-chart-s-data-label-font-specific-typeface-size-bold-styling-emphasis.cs) | `Font`, `Document`, `DocumentBuilder` | Customize chart s data label font specific typeface size bold styling emphasis |
| [define-default-chartdatalabel-options-apply-consistent-...](./define-default-chartdatalabel-options-apply-consistent-font-size-color-across-all.cs) | `DocumentBuilder`, `Series`, `Document` | Define default chartdatalabel options apply consistent font size color across... |
| [enable-automatic-resizing-chart-elements-when-document-...](./enable-automatic-resizing-chart-elements-when-document-page-size-changes-maintain.cs) | `Series`, `Document`, `DocumentBuilder` | Enable automatic resizing chart elements when document page size changes main... |
| [enable-data-label-leader-lines-pie-chart-customize-thei...](./enable-data-label-leader-lines-pie-chart-customize-their-length-better-placement.cs) | `Document`, `DocumentBuilder`, `Series` | Enable data label leader lines pie chart customize their length better placement |
| [existing-docx-file-locate-chart-shape-its-title-replace...](./existing-docx-file-locate-chart-shape-its-title-replace-its-data-source.cs) | `Document`, `Drawing`, `InputDocument` | Existing docx file locate chart shape its title replace its data source |
| [implement-error-handling-catch-exceptions-when-insertin...](./implement-error-handling-catch-exceptions-when-inserting-chart-read-only-document.cs) | `Document`, `DocumentBuilder`, `Series` | Implement error handling catch exceptions when inserting chart read only docu... |
| [insert-chart-table-cell-ensure-it-scales-proportionally...](./insert-chart-table-cell-ensure-it-scales-proportionally-cell-dimensions.cs) | `Document`, `DocumentBuilder`, `CellFormat` | Insert chart table cell ensure it scales proportionally cell dimensions |
| [insert-chart-two-dimensional-array-as-custom-data-sourc...](./insert-chart-two-dimensional-array-as-custom-data-source-mapping-series-categories.cs) | `Document`, `DocumentBuilder`, `Series` | Insert chart two dimensional array as custom data source mapping series categ... |
| [insert-column-chart-new-document-documentbuilder-insert...](./insert-column-chart-new-document-documentbuilder-insertchart-default-data.cs) | `Document`, `DocumentBuilder`, `Drawing` | Insert column chart new document documentbuilder insertchart default data |
| [new-chart-series-set-its-values-via-series-values-prope...](./new-chart-series-set-its-values-via-series-values-property-assign-custom-category.cs) | `Document`, `DocumentBuilder`, `Series` | New chart series set its values via series values property assign custom cate... |
| [programmatically-change-chart-s-plot-area-border-dashed...](./programmatically-change-chart-s-plot-area-border-dashed-line-specific-color-width.cs) | `DocumentBuilder`, `Format`, `Stroke` | Programmatically change chart s plot area border dashed line specific color w... |
| [programmatically-set-chart-s-background-fill-semi-trans...](./programmatically-set-chart-s-background-fill-semi-transparent-color-watermark-effect.cs) | `Document`, `DocumentBuilder`, `Series` | Programmatically set chart s background fill semi transparent color watermark... |
| [remove-specific-series-chart-chartseriescollection-remo...](./remove-specific-series-chart-chartseriescollection-removeat-correct-index.cs) | `Series`, `Document`, `DocumentBuilder` | Remove specific series chart chartseriescollection removeat correct index |
| [retrieve-existing-chart-series-modify-their-data-points...](./retrieve-existing-chart-series-modify-their-data-points-refresh-chart-display.cs) | `ChartYValue`, `ChartXValue`, `YValues` | Retrieve existing chart series modify their data points refresh chart display |
| [retrieve-shape-chart-object-inserted-chart-modify-its-t...](./retrieve-shape-chart-object-inserted-chart-modify-its-title-text-programmatically.cs) | `Document`, `DocumentBuilder`, `Font` | Retrieve shape chart object inserted chart modify its title text programmatic... |
| [set-secondary-y-axis-number-format-currency-two-decimal...](./set-secondary-y-axis-number-format-currency-two-decimal-places-financial-charts.cs) | `Series`, `Document`, `DocumentBuilder` | Set secondary y axis number format currency two decimal places financial charts |
| [update-chart-title-text-toggle-legend-visibility-based-...](./update-chart-title-text-toggle-legend-visibility-based-user-preferences.cs) | `FileNotFoundException`, `InputChart`, `OutputChart` | Update chart title text toggle legend visibility based user preferences |
| [validate-that-all-chart-series-have-matching-category-c...](./validate-that-all-chart-series-have-matching-category-counts-prevent-data.cs) | `Series`, `InvalidOperationException`, `Document` | Validate that all chart series have matching category counts prevent data |
| ... | | *and 1 more files* |

## Category Statistics
- Total examples: 31

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for charts patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
