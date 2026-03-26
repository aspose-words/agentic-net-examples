---
name: linq-reporting
description: C# examples for linq-reporting using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - linq-reporting

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **linq-reporting** category.
This folder contains standalone C# examples for linq-reporting operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **linq-reporting**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using System;` (306/314 files) ← category-specific
- `using Aspose.Words;` (291/314 files)
- `using Aspose.Words.Reporting;` (123/314 files)
- `using System.IO;` (71/314 files)
- `using Aspose.Words.Fields;` (65/314 files)
- `using System.Collections.Generic;` (63/314 files)
- `using System.Data;` (47/314 files)
- `using System.Linq;` (42/314 files)
- `using Aspose.Words.Drawing;` (29/314 files)
- `using Aspose.Words.Saving;` (26/314 files)
- `using Aspose.Words.Tables;` (22/314 files)
- `using Aspose.Words.MailMerging;` (22/314 files)

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
| [access-static-members-system-math-inside-expression-com...](./access-static-members-system-math-inside-expression-compute-rounded-values-financial.cs) | `Document`, `DocumentBuilder`, `MathRoundResultFormatter` | Access static members system math inside expression compute rounded values fi... |
| [activate-reportbuildoptions-inlineerrormessages-capture...](./activate-reportbuildoptions-inlineerrormessages-capture-boolean-success-flag.cs) | `Document`, `ReportingEngine`, `Template` | Activate reportbuildoptions inlineerrormessages capture boolean success flag |
| [add-assembly-attribute-merge-reporting-engine-custom-lo...](./add-assembly-attribute-merge-reporting-engine-custom-logging-assembly-unified-error.cs) | `ReportingEngine`, `Document`, `MyCustomLogging` | Add assembly attribute merge reporting engine custom logging assembly unified... |
| [add-assembly-attribute-required-merging-reporting-engin...](./add-assembly-attribute-required-merging-reporting-engine-custom-authentication-module.cs) | `ReportingEngine` | Add assembly attribute required merging reporting engine custom authenticatio... |
| [add-bookmarks-inside-list-items-ensuring-they-are-not-p...](./add-bookmarks-inside-list-items-ensuring-they-are-not-placed-within-chart-elements.cs) | `Document`, `DocumentBuilder`, `ListFormat` | Add bookmarks inside list items ensuring they are not placed within chart ele... |
| [add-bookmarks-table-rows-hyperlinks-that-navigate-those...](./add-bookmarks-table-rows-hyperlinks-that-navigate-those-bookmarks.cs) | `Document`, `DocumentBuilder`, `Font` | Add bookmarks table rows hyperlinks that navigate those bookmarks |
| [add-custom-security-restrictions-populating-restrictedm...](./add-custom-security-restrictions-populating-restrictedmembers-property-names-that.cs) | `Document`, `DocumentBuilder`, `ReportingEngine` | Add custom security restrictions populating restrictedmembers property names... |
| [add-dynamic-footer-displaying-current-date-page-number-...](./add-dynamic-footer-displaying-current-date-page-number-field-tags.cs) | `Document`, `DocumentBuilder`, `FieldType` | Add dynamic footer displaying current date page number field tags |
| [add-else-clause-after-if-block-display-out-stock-when-s...](./add-else-clause-after-if-block-display-out-stock-when-stock-quantity-equals-zero.cs) | `Document`, `DocumentBuilder`, `FieldType` | Add else clause after if block display out stock when stock quantity equals zero |
| [add-image-expression-tags-docx-template-dynamic-image-i...](./add-image-expression-tags-docx-template-dynamic-image-insertion.cs) | `Document`, `ReportingEngine`, `Template` | Add image expression tags docx template dynamic image insertion |
| [add-optional-display-text-link-tags-defaulting-target-u...](./add-optional-display-text-link-tags-defaulting-target-url-when-omitted.cs) | `Font`, `Document`, `DocumentBuilder` | Add optional display text link tags defaulting target url when omitted |
| [add-required-assembly-attribute-merge-reporting-engine-...](./add-required-assembly-attribute-merge-reporting-engine-another-assembly-deployment.cs) | `ReportingEngine`, `Document`, `DocumentBuilder` | Add required assembly attribute merge reporting engine another assembly deplo... |
| [add-system-datetime-knowntypes-so-templates-can-call-da...](./add-system-datetime-knowntypes-so-templates-can-call-datetime-now-directly.cs) | `Document`, `DocumentBuilder`, `ReportingEngine` | Add system datetime knowntypes so templates can call datetime now directly |
| [aggregate-values-nested-data-bands-expression-tags-sum-...](./aggregate-values-nested-data-bands-expression-tags-sum-function-summary-table.cs) | `Orders`, `Order`, `Customer` | Aggregate values nested data bands expression tags sum function summary table |
| [apply-background-color-table-row-based-status-setbackgr...](./apply-background-color-table-row-based-status-setbackgroundcolor-tag-conditional.cs) | `Color`, `Document`, `DocumentBuilder` | Apply background color table row based status setbackgroundcolor tag conditional |
| [apply-built-any-method-tags-collection-display-tag-list...](./apply-built-any-method-tags-collection-display-tag-list-only-when-tags-exist.cs) | `Document`, `Input`, `Range` | Apply built any method tags collection display tag list only when tags exist |
| [apply-built-count-method-items-where-quantity-0-count-a...](./apply-built-count-method-items-where-quantity-0-count-available-items.cs) | `Collections` | Apply built count method items where quantity 0 count available items |
| [apply-built-firstordefault-extension-method-retrieve-fi...](./apply-built-firstordefault-extension-method-retrieve-first-matching-record-employees.cs) | `Employee`, `Collections` | Apply built firstordefault extension method retrieve first matching record em... |
| [apply-built-min-method-prices-find-lowest-price-display...](./apply-built-min-method-prices-find-lowest-price-display-as-discount-benchmark.cs) | `Document`, `DocumentBuilder`, `Drawing` | Apply built min method prices find lowest price display as discount benchmark |
| [apply-built-where-extension-method-orders-collection-fi...](./apply-built-where-extension-method-orders-collection-filter-records-status-pending.cs) | `Document`, `DocumentBuilder`, `Collections` | Apply built where extension method orders collection filter records status pe... |
| [apply-bulleted-list-template-bullet-point-summary-filte...](./apply-bulleted-list-template-bullet-point-summary-filtered-json-array.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Apply bulleted list template bullet point summary filtered json array |
| [apply-bulleted-list-template-each-group-groupby-produce...](./apply-bulleted-list-template-each-group-groupby-produce-nested-bullet-sections.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Apply bulleted list template each group groupby produce nested bullet sections |
| [apply-bulleted-list-template-each-group-produced-groupb...](./apply-bulleted-list-template-each-group-produced-groupby-nested-bullet-sections.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Apply bulleted list template each group produced groupby nested bullet sections |
| [apply-bulleted-list-template-filtered-xml-node-set-prod...](./apply-bulleted-list-template-filtered-xml-node-set-produce-hierarchical-bullet.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Apply bulleted list template filtered xml node set produce hierarchical bullet |
| [apply-bulleted-list-template-grouped-csv-data-nested-bu...](./apply-bulleted-list-template-grouped-csv-data-nested-bullet-points-each-category.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Apply bulleted list template grouped csv data nested bullet points each category |
| [apply-bulleted-list-template-grouped-json-data-nested-b...](./apply-bulleted-list-template-grouped-json-data-nested-bullet-points-each-category.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Apply bulleted list template grouped json data nested bullet points each cate... |
| [apply-bulleted-list-template-xml-nodes-representing-tas...](./apply-bulleted-list-template-xml-nodes-representing-tasks-bullet-point-list.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Apply bulleted list template xml nodes representing tasks bullet point list |
| [apply-conditional-block-tags-discount-information-only-...](./apply-conditional-block-tags-discount-information-only-when-discount-percentage.cs) | `Document`, `DocumentBuilder`, `DiscountInfo` | Apply conditional block tags discount information only when discount percentage |
| [apply-conditional-formatting-table-rows-evaluating-bool...](./apply-conditional-formatting-table-rows-evaluating-boolean-expression-checking-if-row.cs) | `DocumentBuilder`, `Document`, `Styles` | Apply conditional formatting table rows evaluating boolean expression checkin... |
| [apply-conditional-visibility-sections-based-user-role-i...](./apply-conditional-visibility-sections-based-user-role-if-role-admin-tag.cs) | `Document`, `DocumentBuilder`, `ConditionalVisibility` | Apply conditional visibility sections based user role if role admin tag |
| ... | | *and 284 more files* |

## Category Statistics
- Total examples: 314

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for linq-reporting patterns.


## Command Reference

### Build and Run

Files in this folder are standalone `.cs` examples. Run one example at a time by copying it into a temporary console project as `Program.cs`.

```bash
# Create a temporary console project from the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\linq-reporting\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

### Notes for Agents and Developers

- Treat every `.cs` file in `linq-reporting/` as a full console program, not a snippet.
- Run one file at a time by copying it to `Program.cs`.
- If a sample needs input documents, images, fonts, or data files, place them in the temporary project directory before running.
- See the root [AGENTS.md](../AGENTS.md) for repository-wide prerequisites, project file template, and testing guidance.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
