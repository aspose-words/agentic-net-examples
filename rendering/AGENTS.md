---
name: rendering
description: Verified C# examples for rendering and layout workflows in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Rendering

## Purpose

This folder is a **live, curated example set** for rendering scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free use of documented Aspose.Words rendering APIs for PDF, TIFF, image, SVG, XPS, layout, font, and page-related workflows.

## Non-negotiable conventions

- Always use documented Aspose.Words APIs directly.
- Always create local sample source documents when a task refers to an existing file, folder, stream, template, or input asset.
- Prefer `Document.Save` with the appropriate save options for full-document rendering tasks.
- Prefer `Aspose.Drawing` over `System.Drawing` unless a documented API explicitly requires `System.Drawing`.
- Keep validation narrow and task-specific.
- Do not invent rendering, OpenType, text shaping, or layout helper APIs.

## Recommended workflow selection

- **Rendering workflow**: 70 examples

This category performed best with light primary rules and a narrow task-specific patch only where needed.

## Validation priorities

1. The code must compile and run without manual input.
2. Required sample inputs must be bootstrapped locally inside the example.
3. Requested output files must be created successfully.
4. Validation should focus only on the exact rendering or layout outcome requested by the task.

## File-to-task reference

- `load-a-docx-file-into-a-document-object-using-the-api.cs`
  - Task: Load a DOCX file into a Document object using the API.
  - Workflow: rendering-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `set-imagesaveoptions-desireddpi-to-300-to-produce-high-resolution-tiff-images-for-printing.cs`
  - Task: Set ImageSaveOptions.DesiredDpi to 300 to produce high‑resolution TIFF images for printing.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `configure-imagesaveoptions-pixelformat-as-format1bppindexed-for-black-and-white-tiff-outpu.cs`
  - Task: Configure ImageSaveOptions.PixelFormat as Format1bppIndexed for black‑and‑white TIFF output in the conversion.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `choose-tiffcompression-ccitt4-via-imagesaveoptions-to-apply-lossless-binary-compression-du.cs`
  - Task: Choose TiffCompression.Ccitt4 via ImageSaveOptions to apply lossless binary compression during saving.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `adjust-thresholdforfloydsteinbergdithering-to-150-to-darken-binary-tiff-images-significant.cs`
  - Task: Adjust ThresholdForFloydSteinbergDithering to 150 to darken binary TIFF images significantly.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `assign-fontsettings-fontfolders-to-a-custom-directory-containing-required-truetype-fonts-f.cs`
  - Task: Assign FontSettings.FontFolders to a custom directory containing required TrueType fonts for rendering.
  - Workflow: rendering-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `enable-fontsettings-enableopentypefeatures-to-render-ligatures-and-stylistic-sets-correctl.cs`
  - Task: Enable FontSettings.EnableOpenTypeFeatures to render ligatures and stylistic sets correctly in TIFFs.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `disable-automatic-font-substitution-by-setting-substitutionsettings-enabled-to-false-durin.cs`
  - Task: Disable automatic font substitution by setting SubstitutionSettings.Enabled to false during rendering.
  - Workflow: rendering-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `define-a-fallback-list-using-fontfallbacksettings-predefinedsettings-for-missing-glyphs-du.cs`
  - Task: Define a fallback list using FontFallbackSettings.PredefinedSettings for missing glyphs during rendering.
  - Workflow: rendering-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `set-imagesaveoptions-pagecount-to-limit-tiff-output-to-the-first-five-pages-only.cs`
  - Task: Set ImageSaveOptions.PageCount to limit TIFF output to the first five pages only.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `use-imagesaveoptions-pagesavingcallback-to-customize-file-naming-for-each-generated-tiff-p.cs`
  - Task: Use ImageSaveOptions.PageSavingCallback to customize file naming for each generated TIFF page.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `save-the-document-as-a-multipage-tiff-using-document-save-with-configured-options.cs`
  - Task: Save the document as a multipage TIFF using Document.Save with configured options.
  - Workflow: rendering-workflow
  - Outputs: doc, tiff
  - Selected engine: mcp
- `batch-process-a-folder-of-docx-files-applying-200-dpi-and-ccitt3-compression-to-each-tiff.cs`
  - Task: Batch process a folder of DOCX files, applying 200 DPI and CCITT3 compression to each TIFF.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, tiff
  - Selected engine: mcp
- `create-a-console-application-that-accepts-input-path-dpi-and-compression-type-arguments-fo.cs`
  - Task: Create a console application that accepts input path, DPI, and compression type arguments for TIFF conversion.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `load-documents-from-a-network-share-configure-custom-font-folders-and-output-tiffs-to-a-lo.cs`
  - Task: Load documents from a network share, configure custom font folders, and output TIFFs to a local directory.
  - Workflow: rendering-workflow
  - Outputs: doc, tiff
  - Selected engine: mcp
- `install-truetype-fonts-on-linux-by-copying-ttf-files-and-pointing-fontsettings-to-the-dire.cs`
  - Task: Install TrueType fonts on Linux by copying .TTF files and pointing FontSettings to the directory.
  - Workflow: rendering-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `render-a-doc-file-on-a-headless-linux-server-using-fontsettings-that-reference-the-linux-f.cs`
  - Task: Render a DOC file on a headless Linux server using FontSettings that reference the Linux font folder.
  - Workflow: rendering-workflow
  - Outputs: doc
  - Selected engine: mcp
- `verify-that-opentype-features-such-as-discretionary-ligatures-appear-correctly-in-the-gene.cs`
  - Task: Verify that OpenType features such as discretionary ligatures appear correctly in the generated TIFF.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `subset-embedded-truetype-fonts-during-rendering-to-reduce-the-resulting-tiff-file-size-sig.cs`
  - Task: Subset embedded TrueType fonts during rendering to reduce the resulting TIFF file size significantly.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `apply-format24bpprgb-pixel-format-for-color-tiff-output-from-docx-documents-accurately.cs`
  - Task: Apply Format24bppRgb pixel format for color TIFF output from DOCX documents accurately.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, tiff
  - Selected engine: mcp
- `generate-a-single-page-tiff-from-the-first-page-of-a-multipage-docx-using-imagesaveoptions.cs`
  - Task: Generate a single‑page TIFF from the first page of a multipage DOCX using ImageSaveOptions.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, tiff
  - Selected engine: mcp
- `set-desireddpi-to-72-for-low-resolution-tiff-suitable-for-email-attachments.cs`
  - Task: Set DesiredDpi to 72 for low‑resolution TIFF suitable for email attachments.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `use-format8bppindexed-pixel-format-to-produce-grayscale-tiffs-efficiently-for-archiving.cs`
  - Task: Use Format8bppIndexed pixel format to produce grayscale TIFFs efficiently for archiving.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `apply-tiffcompression-ccitt3-and-verify-file-size-reduction-compared-to-uncompressed-tiff.cs`
  - Task: Apply TiffCompression.Ccitt3 and verify file size reduction compared to uncompressed TIFF significant.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `configure-fontsettings-substitutionsettings-fontfallbacksettings-with-a-custom-xml-hierarc.cs`
  - Task: Configure FontSettings.SubstitutionSettings.FontFallbackSettings with a custom XML hierarchy for fallback fonts definition.
  - Workflow: rendering-workflow
  - Outputs: xml
  - Selected engine: mcp
- `enable-font-substitution-fallback-to-arial-unicode-ms-for-characters-missing-in-primary-fo.cs`
  - Task: Enable font substitution fallback to Arial Unicode MS for characters missing in primary fonts.
  - Workflow: rendering-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `create-a-powershell-script-that-iterates-over-doc-files-sets-dpi-to-250-and-saves-tiffs.cs`
  - Task: Create a PowerShell script that iterates over DOC files, sets DPI to 250, and saves TIFFs.
  - Workflow: rendering-workflow
  - Outputs: doc, tiff
  - Selected engine: mcp
- `implement-parallel-processing-to-batch-convert-doc-files-to-tiff-for-improved-conversion-t.cs`
  - Task: Implement parallel processing to batch convert DOC files to TIFF for improved conversion throughput significantly.
  - Workflow: rendering-workflow
  - Outputs: doc, tiff
  - Selected engine: mcp
- `set-imagesaveoptions-compression-to-tiffcompression-ccitt3-and-pixelformat-to-format1bppin.cs`
  - Task: Set ImageSaveOptions.Compression to TiffCompression.Ccitt3 and PixelFormat to Format1bppIndexed for smallest file size possible.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `use-desireddpi-500-to-generate-archival-quality-tiffs-and-assess-the-impact-on-file-size.cs`
  - Task: Use DesiredDpi 500 to generate archival‑quality TIFFs and assess the impact on file size.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `load-a-docx-set-fontsettings-enableopentypefeatures-false-and-render-to-1bpp-tiff-for-mini.cs`
  - Task: Load a DOCX, set FontSettings.EnableOpenTypeFeatures false, and render to 1bpp TIFF for minimal size.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, tiff
  - Selected engine: mcp
- `configure-fontsettings-to-load-fonts-from-a-usb-drive-for-documents-containing-special-sym.cs`
  - Task: Configure FontSettings to load fonts from a USB drive for documents containing special symbols.
  - Workflow: rendering-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-unit-test-that-asserts-the-tiff-output-contains-the-number-of-pages-as-the-source.cs`
  - Task: Create a unit test that asserts the TIFF output contains the number of pages as the source DOCX.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, tiff
  - Selected engine: mcp
- `set-imagesaveoptions-thresholdforfloydsteinbergdithering-to-90-for-lighter-binary-tiffs-su.cs`
  - Task: Set ImageSaveOptions.ThresholdForFloydSteinbergDithering to 90 for lighter binary TIFFs suitable for OCR preprocessing efficient.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `use-imagesaveoptions-compression-tiffcompression-ccitt4-and-desireddpi-250-for-balanced-qu.cs`
  - Task: Use ImageSaveOptions.Compression = TiffCompression.Ccitt4 and DesiredDpi = 250 for balanced quality black‑and‑white TIFFs.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `load-a-docx-document-and-configure-rendering-options-before-conversion.cs`
  - Task: Load a DOCX document and configure rendering options before conversion.
  - Workflow: rendering-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `create-a-custom-fontsettings-object-pointing-to-a-network-folder-containing-truetype-fonts.cs`
  - Task: Create a custom FontSettings object pointing to a network folder containing TrueType fonts.
  - Workflow: rendering-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `assign-the-fontsettings-instance-to-the-document-to-enable-custom-font-lookup.cs`
  - Task: Assign the FontSettings instance to the Document to enable custom font lookup.
  - Workflow: rendering-workflow
  - Outputs: doc
  - Selected engine: mcp
- `specify-a-default-fallback-font-name-for-missing-truetype-fonts-when-saving-as-pdf.cs`
  - Task: Specify a default fallback font name for missing TrueType fonts when saving as PDF.
  - Workflow: rendering-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `load-font-substitution-rules-from-an-xml-file-and-apply-them-to-the-document.cs`
  - Task: Load font substitution rules from an XML file and apply them to the document.
  - Workflow: rendering-workflow
  - Outputs: doc, xml
  - Selected engine: mcp
- `enable-opentype-typographic-features-to-preserve-ligatures-and-contextual-forms-during-pdf.cs`
  - Task: Enable OpenType typographic features to preserve ligatures and contextual forms during PDF rendering.
  - Workflow: rendering-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `render-a-docx-file-to-pdf-while-embedding-only-the-glyphs-used-from-each-truetype-font.cs`
  - Task: Render a DOCX file to PDF while embedding only the glyphs used from each TrueType font.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, pdf
  - Selected engine: mcp
- `disable-opentype-features-to-generate-a-pdf-for-typographic-comparison.cs`
  - Task: Disable OpenType features to generate a PDF for typographic comparison.
  - Workflow: rendering-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `set-imagesaveoptions-dpix-and-dpiy-to-600-to-produce-high-resolution-tiff-output.cs`
  - Task: Set ImageSaveOptions.DpiX and DpiY to 600 to produce high‑resolution TIFF output.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `set-imagesaveoptions-dpix-and-dpiy-to-300-for-standard-dpi-tiff-conversion.cs`
  - Task: Set ImageSaveOptions.DpiX and DpiY to 300 for standard DPI TIFF conversion.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `apply-a-binarization-threshold-of-150-for-darker-grayscale-tiff-rendering.cs`
  - Task: Apply a binarization threshold of 150 for darker grayscale TIFF rendering.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `apply-a-threshold-of-100-for-lighter-grayscale-tiff-conversion-during-processing.cs`
  - Task: Apply a threshold of 100 for lighter grayscale TIFF conversion during processing.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `configure-imagesaveoptions-pixelformat-to-format1bppindexed-for-1-bit-black-white-tiff-out.cs`
  - Task: Configure ImageSaveOptions.PixelFormat to Format1bppIndexed for 1‑bit black‑white TIFF output generation.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `configure-imagesaveoptions-pixelformat-to-format8bppindexed-and-use-lzw-compression-for-co.cs`
  - Task: Configure ImageSaveOptions.PixelFormat to Format8bppIndexed and use LZW compression for color TIFF.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `set-imagesaveoptions-tiffcompression-to-ccitt4-for-optimal-black-white-tiff-size-reduction.cs`
  - Task: Set ImageSaveOptions.TiffCompression to CCITT4 for optimal black‑white TIFF size reduction.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `set-imagesaveoptions-tiffcompression-to-ccitt3-for-fax-ready-tiff-output-compliant-with-st.cs`
  - Task: Set ImageSaveOptions.TiffCompression to CCITT3 for fax‑ready TIFF output compliant with standards.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `set-imagesaveoptions-compressionlevel-to-maximum-to-reduce-tiff-file-size-without-quality.cs`
  - Task: Set ImageSaveOptions.CompressionLevel to maximum to reduce TIFF file size without quality loss.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `generate-a-multipage-tiff-from-a-docx-file-using-parallel-batch-processing.cs`
  - Task: Generate a multipage TIFF from a DOCX file using parallel batch processing.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, tiff
  - Selected engine: mcp
- `convert-a-batch-of-doc-files-to-tiff-using-1-bit-pixel-format-and-ccitt4-compression.cs`
  - Task: Convert a batch of DOC files to TIFF using 1‑bit pixel format and CCITT4 compression.
  - Workflow: rendering-workflow
  - Outputs: doc, tiff
  - Selected engine: mcp
- `batch-convert-a-folder-of-docx-files-to-tiff-with-shared-imagesaveoptions-settings.cs`
  - Task: Batch convert a folder of DOCX files to TIFF with shared ImageSaveOptions settings.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, tiff
  - Selected engine: mcp
- `create-a-net-service-that-watches-a-folder-and-converts-incoming-docx-files-to-tiff-automa.cs`
  - Task: Create a .NET service that watches a folder and converts incoming DOCX files to TIFF automatically.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, tiff
  - Selected engine: mcp
- `create-a-net-core-console-app-that-monitors-a-directory-and-converts-new-docx-files-to-tif.cs`
  - Task: Create a .NET Core console app that monitors a directory and converts new DOCX files to TIFF.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, tiff
  - Selected engine: mcp
- `write-a-powershell-script-to-batch-convert-docx-files-to-multipage-tiff-with-custom-dpi.cs`
  - Task: Write a PowerShell script to batch convert DOCX files to multipage TIFF with custom DPI.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, tiff
  - Selected engine: mcp
- `write-a-batch-script-that-converts-all-docx-files-in-a-folder-to-tiff-using-predefined-com.cs`
  - Task: Write a batch script that converts all DOCX files in a folder to TIFF using predefined compression.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, tiff
  - Selected engine: mcp
- `adjust-imagesaveoptions-thresholdforfloydsteinbergdithering-to-200-for-darker-tiff-output.cs`
  - Task: Adjust ImageSaveOptions.ThresholdForFloydSteinbergDithering to 200 for darker TIFF output during rendering.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `adjust-imagesaveoptions-thresholdforfloydsteinbergdithering-to-180-for-moderately-dark-tif.cs`
  - Task: Adjust ImageSaveOptions.ThresholdForFloydSteinbergDithering to 180 for moderately dark TIFF images during conversion.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `set-imagesaveoptions-dpix-to-72-and-dpiy-to-144-to-create-vertically-stretched-tiff.cs`
  - Task: Set ImageSaveOptions.DpiX to 72 and DpiY to 144 to create vertically stretched TIFF.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `set-imagesaveoptions-dpix-to-72-and-dpiy-to-300-for-asymmetric-resolution-tiff.cs`
  - Task: Set ImageSaveOptions.DpiX to 72 and DpiY to 300 for asymmetric resolution TIFF.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `split-a-generated-multipage-tiff-into-individual-page-images-programmatically-after-conver.cs`
  - Task: Split a generated multipage TIFF into individual page images programmatically after conversion.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `create-a-unit-test-that-verifies-tiff-output-matches-expected-pixel-format-and-compression.cs`
  - Task: Create a unit test that verifies TIFF output matches expected pixel format and compression.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `create-a-unit-test-that-ensures-1-bit-tiff-files-contain-exactly-two-colors.cs`
  - Task: Create a unit test that ensures 1‑bit TIFF files contain exactly two colors.
  - Workflow: rendering-workflow
  - Outputs: tiff
  - Selected engine: mcp
- `extract-embedded-truetype-fonts-from-a-pdf-to-verify-correct-subsetting-after-rendering.cs`
  - Task: Extract embedded TrueType fonts from a PDF to verify correct subsetting after rendering.
  - Workflow: rendering-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `render-a-docx-to-pdf-with-full-font-embedding-and-then-extract-fonts-to-confirm-subsetting.cs`
  - Task: Render a DOCX to PDF with full font embedding and then extract fonts to confirm subsetting disabled.
  - Workflow: rendering-workflow
  - Outputs: docx, doc, pdf
  - Selected engine: mcp
- `configure-fontsettings-substitutionsettings-to-map-missing-arial-to-liberation-sans-on-lin.cs`
  - Task: Configure FontSettings.SubstitutionSettings to map missing Arial to Liberation Sans on Linux platforms.
  - Workflow: rendering-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `configure-fontsettings-to-prioritize-fonts-from-a-user-specified-directory-over-system-fon.cs`
  - Task: Configure FontSettings to prioritize fonts from a user‑specified directory over system fonts during rendering.
  - Workflow: rendering-workflow
  - Outputs: pdf
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **System.Drawing inspection in verifier**
  - Symptom: Build failures around `Bitmap`, `Graphics`, `Image`, or `FrameDimension`.
  - Preferred fix: Prefer `Aspose.Drawing` where possible, and avoid System.Drawing-based TIFF inspection unless a documented API explicitly requires it.
- **Invented OpenType or shaping APIs**
  - Symptom: Build failures around `EnableOpenTypeFeatures`, `HarfBuzzTextShaperFactory`, `BasicTextShaperFactory`, or similar names.
  - Preferred fix: Use only documented font and rendering APIs available in the current Aspose.Words package version.
- **Overcomplicated layout traversal**
  - Symptom: Build or runtime failures caused by unnecessary layout or renderer APIs for simple save/render tasks.
  - Preferred fix: Use `Document.Save` and simple output validation unless the task explicitly requires page or node renderer behavior.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`
- Helper drawing package: `Aspose.Drawing.Common`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.Drawing.Common
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\rendering\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve exact file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the verified winner from the latest batch report rather than a merely compiling draft.
- Bootstrap file-based inputs locally instead of depending on machine-specific paths.
