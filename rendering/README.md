# Rendering Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **Rendering** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Rendering**
- Slug: **rendering**
- Total examples: **70**
- Workflow examples: **70 / 70** use the standard rendering workflow

## Category rules that shaped these examples

- Use native Aspose.Words APIs directly.
- Create local sample source documents when a task refers to an existing file, folder, stream, template, or input asset.
- Do not assume external files or folders already exist.
- Prefer `Document.Save` with the appropriate save options for rendering tasks.
- Prefer `Aspose.Drawing` over `System.Drawing` unless a documented API explicitly requires `System.Drawing`.
- Keep validation narrow and task-specific.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`
- Aspose.Drawing.Common

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.Drawing.Common

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\rendering\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.Drawing.Common

# PowerShell example
Copy-Item ..\rendering\load-a-docx-file-into-a-document-object-using-the-api.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-file-into-a-document-object-using-the-api.cs` | Load a DOCX file into a Document object using the API. | rendering-workflow | docx, doc | mcp |
| 2 | `set-imagesaveoptions-desireddpi-to-300-to-produce-high-resolution-tiff-images-for-printing.cs` | Set ImageSaveOptions.DesiredDpi to 300 to produce high‑resolution TIFF images for printing. | rendering-workflow | tiff | mcp |
| 3 | `configure-imagesaveoptions-pixelformat-as-format1bppindexed-for-black-and-white-tiff-outpu.cs` | Configure ImageSaveOptions.PixelFormat as Format1bppIndexed for black‑and‑white TIFF output in the conversion. | rendering-workflow | tiff | mcp |
| 4 | `choose-tiffcompression-ccitt4-via-imagesaveoptions-to-apply-lossless-binary-compression-du.cs` | Choose TiffCompression.Ccitt4 via ImageSaveOptions to apply lossless binary compression during saving. | rendering-workflow | tiff | mcp |
| 5 | `adjust-thresholdforfloydsteinbergdithering-to-150-to-darken-binary-tiff-images-significant.cs` | Adjust ThresholdForFloydSteinbergDithering to 150 to darken binary TIFF images significantly. | rendering-workflow | tiff | mcp |
| 6 | `assign-fontsettings-fontfolders-to-a-custom-directory-containing-required-truetype-fonts-f.cs` | Assign FontSettings.FontFolders to a custom directory containing required TrueType fonts for rendering. | rendering-workflow | pdf | mcp |
| 7 | `enable-fontsettings-enableopentypefeatures-to-render-ligatures-and-stylistic-sets-correctl.cs` | Enable FontSettings.EnableOpenTypeFeatures to render ligatures and stylistic sets correctly in TIFFs. | rendering-workflow | tiff | mcp |
| 8 | `disable-automatic-font-substitution-by-setting-substitutionsettings-enabled-to-false-durin.cs` | Disable automatic font substitution by setting SubstitutionSettings.Enabled to false during rendering. | rendering-workflow | pdf | mcp |
| 9 | `define-a-fallback-list-using-fontfallbacksettings-predefinedsettings-for-missing-glyphs-du.cs` | Define a fallback list using FontFallbackSettings.PredefinedSettings for missing glyphs during rendering. | rendering-workflow | pdf | mcp |
| 10 | `set-imagesaveoptions-pagecount-to-limit-tiff-output-to-the-first-five-pages-only.cs` | Set ImageSaveOptions.PageCount to limit TIFF output to the first five pages only. | rendering-workflow | tiff | mcp |
| 11 | `use-imagesaveoptions-pagesavingcallback-to-customize-file-naming-for-each-generated-tiff-p.cs` | Use ImageSaveOptions.PageSavingCallback to customize file naming for each generated TIFF page. | rendering-workflow | tiff | mcp |
| 12 | `save-the-document-as-a-multipage-tiff-using-document-save-with-configured-options.cs` | Save the document as a multipage TIFF using Document.Save with configured options. | rendering-workflow | doc, tiff | mcp |
| 13 | `batch-process-a-folder-of-docx-files-applying-200-dpi-and-ccitt3-compression-to-each-tiff.cs` | Batch process a folder of DOCX files, applying 200 DPI and CCITT3 compression to each TIFF. | rendering-workflow | docx, doc, tiff | mcp |
| 14 | `create-a-console-application-that-accepts-input-path-dpi-and-compression-type-arguments-fo.cs` | Create a console application that accepts input path, DPI, and compression type arguments for TIFF conversion. | rendering-workflow | tiff | mcp |
| 15 | `load-documents-from-a-network-share-configure-custom-font-folders-and-output-tiffs-to-a-lo.cs` | Load documents from a network share, configure custom font folders, and output TIFFs to a local directory. | rendering-workflow | doc, tiff | mcp |
| 16 | `install-truetype-fonts-on-linux-by-copying-ttf-files-and-pointing-fontsettings-to-the-dire.cs` | Install TrueType fonts on Linux by copying .TTF files and pointing FontSettings to the directory. | rendering-workflow | pdf | mcp |
| 17 | `render-a-doc-file-on-a-headless-linux-server-using-fontsettings-that-reference-the-linux-f.cs` | Render a DOC file on a headless Linux server using FontSettings that reference the Linux font folder. | rendering-workflow | doc | mcp |
| 18 | `verify-that-opentype-features-such-as-discretionary-ligatures-appear-correctly-in-the-gene.cs` | Verify that OpenType features such as discretionary ligatures appear correctly in the generated TIFF. | rendering-workflow | tiff | mcp |
| 19 | `subset-embedded-truetype-fonts-during-rendering-to-reduce-the-resulting-tiff-file-size-sig.cs` | Subset embedded TrueType fonts during rendering to reduce the resulting TIFF file size significantly. | rendering-workflow | tiff | mcp |
| 20 | `apply-format24bpprgb-pixel-format-for-color-tiff-output-from-docx-documents-accurately.cs` | Apply Format24bppRgb pixel format for color TIFF output from DOCX documents accurately. | rendering-workflow | docx, doc, tiff | mcp |
| 21 | `generate-a-single-page-tiff-from-the-first-page-of-a-multipage-docx-using-imagesaveoptions.cs` | Generate a single‑page TIFF from the first page of a multipage DOCX using ImageSaveOptions. | rendering-workflow | docx, doc, tiff | mcp |
| 22 | `set-desireddpi-to-72-for-low-resolution-tiff-suitable-for-email-attachments.cs` | Set DesiredDpi to 72 for low‑resolution TIFF suitable for email attachments. | rendering-workflow | tiff | mcp |
| 23 | `use-format8bppindexed-pixel-format-to-produce-grayscale-tiffs-efficiently-for-archiving.cs` | Use Format8bppIndexed pixel format to produce grayscale TIFFs efficiently for archiving. | rendering-workflow | tiff | mcp |
| 24 | `apply-tiffcompression-ccitt3-and-verify-file-size-reduction-compared-to-uncompressed-tiff.cs` | Apply TiffCompression.Ccitt3 and verify file size reduction compared to uncompressed TIFF significant. | rendering-workflow | tiff | mcp |
| 25 | `configure-fontsettings-substitutionsettings-fontfallbacksettings-with-a-custom-xml-hierarc.cs` | Configure FontSettings.SubstitutionSettings.FontFallbackSettings with a custom XML hierarchy for fallback fonts definition. | rendering-workflow | xml | mcp |
| 26 | `enable-font-substitution-fallback-to-arial-unicode-ms-for-characters-missing-in-primary-fo.cs` | Enable font substitution fallback to Arial Unicode MS for characters missing in primary fonts. | rendering-workflow | pdf | mcp |
| 27 | `create-a-powershell-script-that-iterates-over-doc-files-sets-dpi-to-250-and-saves-tiffs.cs` | Create a PowerShell script that iterates over DOC files, sets DPI to 250, and saves TIFFs. | rendering-workflow | doc, tiff | mcp |
| 28 | `implement-parallel-processing-to-batch-convert-doc-files-to-tiff-for-improved-conversion-t.cs` | Implement parallel processing to batch convert DOC files to TIFF for improved conversion throughput significantly. | rendering-workflow | doc, tiff | mcp |
| 29 | `set-imagesaveoptions-compression-to-tiffcompression-ccitt3-and-pixelformat-to-format1bppin.cs` | Set ImageSaveOptions.Compression to TiffCompression.Ccitt3 and PixelFormat to Format1bppIndexed for smallest file size possible. | rendering-workflow | tiff | mcp |
| 30 | `use-desireddpi-500-to-generate-archival-quality-tiffs-and-assess-the-impact-on-file-size.cs` | Use DesiredDpi 500 to generate archival‑quality TIFFs and assess the impact on file size. | rendering-workflow | tiff | mcp |
| 31 | `load-a-docx-set-fontsettings-enableopentypefeatures-false-and-render-to-1bpp-tiff-for-mini.cs` | Load a DOCX, set FontSettings.EnableOpenTypeFeatures false, and render to 1bpp TIFF for minimal size. | rendering-workflow | docx, doc, tiff | mcp |
| 32 | `configure-fontsettings-to-load-fonts-from-a-usb-drive-for-documents-containing-special-sym.cs` | Configure FontSettings to load fonts from a USB drive for documents containing special symbols. | rendering-workflow | doc | mcp |
| 33 | `create-a-unit-test-that-asserts-the-tiff-output-contains-the-number-of-pages-as-the-source.cs` | Create a unit test that asserts the TIFF output contains the number of pages as the source DOCX. | rendering-workflow | docx, doc, tiff | mcp |
| 34 | `set-imagesaveoptions-thresholdforfloydsteinbergdithering-to-90-for-lighter-binary-tiffs-su.cs` | Set ImageSaveOptions.ThresholdForFloydSteinbergDithering to 90 for lighter binary TIFFs suitable for OCR preprocessing efficient. | rendering-workflow | tiff | mcp |
| 35 | `use-imagesaveoptions-compression-tiffcompression-ccitt4-and-desireddpi-250-for-balanced-qu.cs` | Use ImageSaveOptions.Compression = TiffCompression.Ccitt4 and DesiredDpi = 250 for balanced quality black‑and‑white TIFFs. | rendering-workflow | tiff | mcp |
| 36 | `load-a-docx-document-and-configure-rendering-options-before-conversion.cs` | Load a DOCX document and configure rendering options before conversion. | rendering-workflow | docx, doc | mcp |
| 37 | `create-a-custom-fontsettings-object-pointing-to-a-network-folder-containing-truetype-fonts.cs` | Create a custom FontSettings object pointing to a network folder containing TrueType fonts. | rendering-workflow | pdf | mcp |
| 38 | `assign-the-fontsettings-instance-to-the-document-to-enable-custom-font-lookup.cs` | Assign the FontSettings instance to the Document to enable custom font lookup. | rendering-workflow | doc | mcp |
| 39 | `specify-a-default-fallback-font-name-for-missing-truetype-fonts-when-saving-as-pdf.cs` | Specify a default fallback font name for missing TrueType fonts when saving as PDF. | rendering-workflow | pdf | mcp |
| 40 | `load-font-substitution-rules-from-an-xml-file-and-apply-them-to-the-document.cs` | Load font substitution rules from an XML file and apply them to the document. | rendering-workflow | doc, xml | mcp |
| 41 | `enable-opentype-typographic-features-to-preserve-ligatures-and-contextual-forms-during-pdf.cs` | Enable OpenType typographic features to preserve ligatures and contextual forms during PDF rendering. | rendering-workflow | pdf | mcp |
| 42 | `render-a-docx-file-to-pdf-while-embedding-only-the-glyphs-used-from-each-truetype-font.cs` | Render a DOCX file to PDF while embedding only the glyphs used from each TrueType font. | rendering-workflow | docx, doc, pdf | mcp |
| 43 | `disable-opentype-features-to-generate-a-pdf-for-typographic-comparison.cs` | Disable OpenType features to generate a PDF for typographic comparison. | rendering-workflow | pdf | mcp |
| 44 | `set-imagesaveoptions-dpix-and-dpiy-to-600-to-produce-high-resolution-tiff-output.cs` | Set ImageSaveOptions.DpiX and DpiY to 600 to produce high‑resolution TIFF output. | rendering-workflow | tiff | mcp |
| 45 | `set-imagesaveoptions-dpix-and-dpiy-to-300-for-standard-dpi-tiff-conversion.cs` | Set ImageSaveOptions.DpiX and DpiY to 300 for standard DPI TIFF conversion. | rendering-workflow | tiff | mcp |
| 46 | `apply-a-binarization-threshold-of-150-for-darker-grayscale-tiff-rendering.cs` | Apply a binarization threshold of 150 for darker grayscale TIFF rendering. | rendering-workflow | tiff | mcp |
| 47 | `apply-a-threshold-of-100-for-lighter-grayscale-tiff-conversion-during-processing.cs` | Apply a threshold of 100 for lighter grayscale TIFF conversion during processing. | rendering-workflow | tiff | mcp |
| 48 | `configure-imagesaveoptions-pixelformat-to-format1bppindexed-for-1-bit-black-white-tiff-out.cs` | Configure ImageSaveOptions.PixelFormat to Format1bppIndexed for 1‑bit black‑white TIFF output generation. | rendering-workflow | tiff | mcp |
| 49 | `configure-imagesaveoptions-pixelformat-to-format8bppindexed-and-use-lzw-compression-for-co.cs` | Configure ImageSaveOptions.PixelFormat to Format8bppIndexed and use LZW compression for color TIFF. | rendering-workflow | tiff | mcp |
| 50 | `set-imagesaveoptions-tiffcompression-to-ccitt4-for-optimal-black-white-tiff-size-reduction.cs` | Set ImageSaveOptions.TiffCompression to CCITT4 for optimal black‑white TIFF size reduction. | rendering-workflow | tiff | mcp |
| 51 | `set-imagesaveoptions-tiffcompression-to-ccitt3-for-fax-ready-tiff-output-compliant-with-st.cs` | Set ImageSaveOptions.TiffCompression to CCITT3 for fax‑ready TIFF output compliant with standards. | rendering-workflow | tiff | mcp |
| 52 | `set-imagesaveoptions-compressionlevel-to-maximum-to-reduce-tiff-file-size-without-quality.cs` | Set ImageSaveOptions.CompressionLevel to maximum to reduce TIFF file size without quality loss. | rendering-workflow | tiff | mcp |
| 53 | `generate-a-multipage-tiff-from-a-docx-file-using-parallel-batch-processing.cs` | Generate a multipage TIFF from a DOCX file using parallel batch processing. | rendering-workflow | docx, doc, tiff | mcp |
| 54 | `convert-a-batch-of-doc-files-to-tiff-using-1-bit-pixel-format-and-ccitt4-compression.cs` | Convert a batch of DOC files to TIFF using 1‑bit pixel format and CCITT4 compression. | rendering-workflow | doc, tiff | mcp |
| 55 | `batch-convert-a-folder-of-docx-files-to-tiff-with-shared-imagesaveoptions-settings.cs` | Batch convert a folder of DOCX files to TIFF with shared ImageSaveOptions settings. | rendering-workflow | docx, doc, tiff | mcp |
| 56 | `create-a-net-service-that-watches-a-folder-and-converts-incoming-docx-files-to-tiff-automa.cs` | Create a .NET service that watches a folder and converts incoming DOCX files to TIFF automatically. | rendering-workflow | docx, doc, tiff | mcp |
| 57 | `create-a-net-core-console-app-that-monitors-a-directory-and-converts-new-docx-files-to-tif.cs` | Create a .NET Core console app that monitors a directory and converts new DOCX files to TIFF. | rendering-workflow | docx, doc, tiff | mcp |
| 58 | `write-a-powershell-script-to-batch-convert-docx-files-to-multipage-tiff-with-custom-dpi.cs` | Write a PowerShell script to batch convert DOCX files to multipage TIFF with custom DPI. | rendering-workflow | docx, doc, tiff | mcp |
| 59 | `write-a-batch-script-that-converts-all-docx-files-in-a-folder-to-tiff-using-predefined-com.cs` | Write a batch script that converts all DOCX files in a folder to TIFF using predefined compression. | rendering-workflow | docx, doc, tiff | mcp |
| 60 | `adjust-imagesaveoptions-thresholdforfloydsteinbergdithering-to-200-for-darker-tiff-output.cs` | Adjust ImageSaveOptions.ThresholdForFloydSteinbergDithering to 200 for darker TIFF output during rendering. | rendering-workflow | tiff | mcp |
| 61 | `adjust-imagesaveoptions-thresholdforfloydsteinbergdithering-to-180-for-moderately-dark-tif.cs` | Adjust ImageSaveOptions.ThresholdForFloydSteinbergDithering to 180 for moderately dark TIFF images during conversion. | rendering-workflow | tiff | mcp |
| 62 | `set-imagesaveoptions-dpix-to-72-and-dpiy-to-144-to-create-vertically-stretched-tiff.cs` | Set ImageSaveOptions.DpiX to 72 and DpiY to 144 to create vertically stretched TIFF. | rendering-workflow | tiff | mcp |
| 63 | `set-imagesaveoptions-dpix-to-72-and-dpiy-to-300-for-asymmetric-resolution-tiff.cs` | Set ImageSaveOptions.DpiX to 72 and DpiY to 300 for asymmetric resolution TIFF. | rendering-workflow | tiff | mcp |
| 64 | `split-a-generated-multipage-tiff-into-individual-page-images-programmatically-after-conver.cs` | Split a generated multipage TIFF into individual page images programmatically after conversion. | rendering-workflow | tiff | mcp |
| 65 | `create-a-unit-test-that-verifies-tiff-output-matches-expected-pixel-format-and-compression.cs` | Create a unit test that verifies TIFF output matches expected pixel format and compression. | rendering-workflow | tiff | mcp |
| 66 | `create-a-unit-test-that-ensures-1-bit-tiff-files-contain-exactly-two-colors.cs` | Create a unit test that ensures 1‑bit TIFF files contain exactly two colors. | rendering-workflow | tiff | mcp |
| 67 | `extract-embedded-truetype-fonts-from-a-pdf-to-verify-correct-subsetting-after-rendering.cs` | Extract embedded TrueType fonts from a PDF to verify correct subsetting after rendering. | rendering-workflow | pdf | mcp |
| 68 | `render-a-docx-to-pdf-with-full-font-embedding-and-then-extract-fonts-to-confirm-subsetting.cs` | Render a DOCX to PDF with full font embedding and then extract fonts to confirm subsetting disabled. | rendering-workflow | docx, doc, pdf | mcp |
| 69 | `configure-fontsettings-substitutionsettings-to-map-missing-arial-to-liberation-sans-on-lin.cs` | Configure FontSettings.SubstitutionSettings to map missing Arial to Liberation Sans on Linux platforms. | rendering-workflow | pdf | mcp |
| 70 | `configure-fontsettings-to-prioritize-fonts-from-a-user-specified-directory-over-system-fon.cs` | Configure FontSettings to prioritize fonts from a user‑specified directory over system fonts during rendering. | rendering-workflow | pdf | mcp |

## Notes for maintainers

- The selected file for each task is the verified winner recorded in the batch run.
- This category performed well with light rules plus a narrow task-specific fix for PDF embedded font verification.
- Preserve exact file-to-task traceability when updating the category.
- Bootstrap all sample input files locally inside the example when the task refers to an existing asset.
