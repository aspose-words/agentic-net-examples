---
name: rendering
description: C# examples for rendering using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - rendering

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **rendering** category.
This folder contains standalone C# examples for rendering operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **rendering**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (63/63 files) ← category-specific
- `using Aspose.Words.Saving;` (50/63 files)
- `using System;` (49/63 files)
- `using System.IO;` (25/63 files)
- `using Aspose.Words.Fonts;` (18/63 files)
- `using System.Drawing;` (4/63 files)
- `using Aspose.Words.Loading;` (3/63 files)
- `using Aspose.Words.Drawing;` (2/63 files)
- `using System.Linq;` (2/63 files)
- `using System.Threading.Tasks;` (2/63 files)
- `using System.Threading;` (2/63 files)
- `using System.Text;` (1/63 files)

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
| [adjust-imagesaveoptions-thresholdforfloydsteinbergdithe...](./adjust-imagesaveoptions-thresholdforfloydsteinbergdithering-180-moderately-dark-tiff.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Adjust imagesaveoptions thresholdforfloydsteinbergdithering 180 moderately da... |
| [adjust-imagesaveoptions-thresholdforfloydsteinbergdithe...](./adjust-imagesaveoptions-thresholdforfloydsteinbergdithering-200-darker-tiff-output.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Adjust imagesaveoptions thresholdforfloydsteinbergdithering 200 darker tiff o... |
| [adjust-thresholdforfloydsteinbergdithering-150-darken-b...](./adjust-thresholdforfloydsteinbergdithering-150-darken-binary-tiff-images-significantly.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Adjust thresholdforfloydsteinbergdithering 150 darken binary tiff images sign... |
| [apply-binarization-threshold-150-darker-grayscale-tiff-...](./apply-binarization-threshold-150-darker-grayscale-tiff-rendering.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Apply binarization threshold 150 darker grayscale tiff rendering |
| [apply-format24bpprgb-pixel-format-color-tiff-output-doc...](./apply-format24bpprgb-pixel-format-color-tiff-output-docx-documents-accurately.cs) | `Document`, `ImageSaveOptions`, `Aspose` | Apply format24bpprgb pixel format color tiff output docx documents accurately |
| [apply-threshold-100-lighter-grayscale-tiff-conversion-d...](./apply-threshold-100-lighter-grayscale-tiff-conversion-during-processing.cs) | `Document`, `ImageSaveOptions`, `Aspose` | Apply threshold 100 lighter grayscale tiff conversion during processing |
| [apply-tiffcompression-ccitt3-verify-file-size-reduction...](./apply-tiffcompression-ccitt3-verify-file-size-reduction-compared-uncompressed-tiff.cs) | `ImageSaveOptions`, `Aspose`, `Document` | Apply tiffcompression ccitt3 verify file size reduction compared uncompressed... |
| [assign-fontsettings-fontfolders-custom-directory-contai...](./assign-fontsettings-fontfolders-custom-directory-containing-required-truetype-fonts.cs) | `Document`, `Aspose`, `FontSettings` | Assign fontsettings fontfolders custom directory containing required truetype... |
| [assign-fontsettings-instance-document-enable-custom-fon...](./assign-fontsettings-instance-document-enable-custom-font-lookup.cs) | `Document`, `FontSettings`, `Aspose` | Assign fontsettings instance document enable custom font lookup |
| [batch-convert-folder-docx-files-tiff-shared-imagesaveop...](./batch-convert-folder-docx-files-tiff-shared-imagesaveoptions-settings.cs) | `ImageSaveOptions`, `Document`, `Aspose` | Batch convert folder docx files tiff shared imagesaveoptions settings |
| [batch-process-folder-docx-files-applying-200-dpi-ccitt3...](./batch-process-folder-docx-files-applying-200-dpi-ccitt3-compression-each-tiff.cs) | `Document`, `ImageSaveOptions`, `Aspose` | Batch process folder docx files applying 200 dpi ccitt3 compression each tiff |
| [choose-tiffcompression-ccitt4-via-imagesaveoptions-appl...](./choose-tiffcompression-ccitt4-via-imagesaveoptions-apply-lossless-binary-compression.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Choose tiffcompression ccitt4 via imagesaveoptions apply lossless binary comp... |
| [configure-fontsettings-fonts-usb-drive-documents-contai...](./configure-fontsettings-fonts-usb-drive-documents-containing-special-symbols.cs) | `LoadOptions`, `Aspose`, `FontSettings` | Configure fontsettings fonts usb drive documents containing special symbols |
| [configure-fontsettings-prioritize-fonts-user-specified-...](./configure-fontsettings-prioritize-fonts-user-specified-directory-over-system-fonts.cs) | `FontSettings`, `Aspose`, `FolderFontSource` | Configure fontsettings prioritize fonts user specified directory over system... |
| [configure-fontsettings-substitutionsettings-fontfallbac...](./configure-fontsettings-substitutionsettings-fontfallbacksettings-custom-xml-hierarchy.cs) | `Document`, `FontSettings`, `System` | Configure fontsettings substitutionsettings fontfallbacksettings custom xml h... |
| [configure-fontsettings-substitutionsettings-map-missing...](./configure-fontsettings-substitutionsettings-map-missing-arial-liberation-sans-linux.cs) | `Document`, `FontSettings`, `Aspose` | Configure fontsettings substitutionsettings map missing arial liberation sans... |
| [configure-imagesaveoptions-pixelformat-as-format1bppind...](./configure-imagesaveoptions-pixelformat-as-format1bppindexed-black-white-tiff-output.cs) | `Document`, `ImageSaveOptions`, `Aspose` | Configure imagesaveoptions pixelformat as format1bppindexed black white tiff... |
| [configure-imagesaveoptions-pixelformat-format1bppindexe...](./configure-imagesaveoptions-pixelformat-format1bppindexed-1-bit-black-white-tiff.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Configure imagesaveoptions pixelformat format1bppindexed 1 bit black white tiff |
| [console-application-that-accepts-input-path-dpi-compres...](./console-application-that-accepts-input-path-dpi-compression-type-arguments-tiff.cs) | `Document`, `ImageSaveOptions`, `Aspose` | Console application that accepts input path dpi compression type arguments tiff |
| [convert-batch-doc-files-tiff-1-bit-pixel-format-ccitt4-...](./convert-batch-doc-files-tiff-1-bit-pixel-format-ccitt4-compression.cs) | `Document`, `ImageSaveOptions`, `Aspose` | Convert batch doc files tiff 1 bit pixel format ccitt4 compression |
| [custom-fontsettings-object-pointing-network-folder-cont...](./custom-fontsettings-object-pointing-network-folder-containing-truetype-fonts.cs) | `Aspose`, `FontSettings`, `Document` | Custom fontsettings object pointing network folder containing truetype fonts |
| [define-fallback-list-fontfallbacksettings-predefinedset...](./define-fallback-list-fontfallbacksettings-predefinedsettings-missing-glyphs-during.cs) | `Document`, `FontSettings`, `DocumentBuilder` | Define fallback list fontfallbacksettings predefinedsettings missing glyphs d... |
| [desireddpi-500-archival-quality-tiffs-assess-impact-fil...](./desireddpi-500-archival-quality-tiffs-assess-impact-file-size.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Desireddpi 500 archival quality tiffs assess impact file size |
| [disable-automatic-font-substitution-setting-substitutio...](./disable-automatic-font-substitution-setting-substitutionsettings-enabled-false-during.cs) | `SubstitutionSettings`, `Document`, `FontSettings` | Disable automatic font substitution setting substitutionsettings enabled fals... |
| [disable-opentype-features-pdf-typographic-comparison](./disable-opentype-features-pdf-typographic-comparison.cs) | `Aspose`, `Document`, `PdfSaveOptions` | Disable opentype features pdf typographic comparison |
| [document-as-multipage-tiff-document-configured-options](./document-as-multipage-tiff-document-configured-options.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Document as multipage tiff document configured options |
| [documents-network-share-configure-custom-font-folders-o...](./documents-network-share-configure-custom-font-folders-output-tiffs-local-directory.cs) | `Aspose`, `Document`, `ImageSaveOptions` | Documents network share configure custom font folders output tiffs local dire... |
| [docx-document-configure-rendering-options-before-conver...](./docx-document-configure-rendering-options-before-conversion.cs) | `LayoutOptions`, `Aspose`, `Document` | Docx document configure rendering options before conversion |
| [docx-file-document-object-api](./docx-file-document-object-api.cs) | `Document`, `Aspose`, `Sample` | Docx file document object api |
| [docx-set-fontsettings-enableopentypefeatures-false-rend...](./docx-set-fontsettings-enableopentypefeatures-false-render-1bpp-tiff-minimal-size.cs) | `Aspose`, `Document`, `ImageSaveOptions` | Docx set fontsettings enableopentypefeatures false render 1bpp tiff minimal size |
| ... | | *and 33 more files* |

## Category Statistics
- Total examples: 63

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for rendering patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082553`
<!-- AUTOGENERATED:END -->
