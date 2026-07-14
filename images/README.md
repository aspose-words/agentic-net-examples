# Images Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Images category. Each file is a standalone console example selected from the verified 26.5.0 run.

## Snapshot

- Category: Images
- Slug: images
- Total examples: 120
- Publish-ready successful examples: 120 / 120
- Source run: 20260619_131835_59df5f
- General Image Workflow examples: 4
- Image Conversion And Manipulation examples: 2
- Image Extraction examples: 109
- Image Insertion And Shape examples: 4
- Input Bootstrap examples: 1

## Category rules that shaped these examples

- Do not use System.Drawing in this category.
- Do not download images from the internet.
- Do not assume image files, folders, or source documents already exist.
- Do not invent unsupported image APIs.
- Create deterministic local sample images with Aspose.Drawing before using them.
- Use DocumentBuilder.InsertImage for standard insertion and Shape.ImageData for shape-based scenarios.
- Enumerate Shape nodes and check HasImage before accessing ImageData.
- Use local deterministic output names and validate that requested output files are written.
- Initialize all non-nullable reference type properties to avoid CS8618 warnings.
- Avoid CS8600, CS8602, and CS8604 by guarding maybe-null values before dereference or assignment.
- Declare nullable locals when a value may legitimately be null and null-check before use.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words 26.5.0
- Aspose.Drawing.Common 25.11.0
- Newtonsoft.Json

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0
dotnet add package Aspose.Drawing.Common --version 25.11.0
dotnet add package Newtonsoft.Json

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\images\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `images/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0
dotnet add package Aspose.Drawing.Common --version 25.11.0
dotnet add package Newtonsoft.Json

# PowerShell example
Copy-Item ..\images\load-a-docx-file-and-extract-all-embedded-images-into-a-specified-output-folder.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-file-and-extract-all-embedded-images-into-a-specified-output-folder.cs` | Load a DOCX file and extract all embedded images into a specified output folder. | Image Extraction | docx | mcp |
| 2 | `convert-each-extracted-png-image-to-jpeg-format-while-preserving-its-original-dimensions.cs` | Convert each extracted PNG image to JPEG format while preserving its original dimensions. | Image Extraction | png, jpeg | mcp |
| 3 | `batch-process-a-folder-of-doc-files-extracting-images-and-saving-them-as-bmp-files.cs` | Batch process a folder of DOC files, extracting images and saving them as BMP files. | Image Extraction | doc, bmp | mcp |
| 4 | `resize-all-jpeg-images-extracted-from-a-word-document-to-a-maximum-width-of-800-pixels.cs` | Resize all JPEG images extracted from a Word document to a maximum width of 800 pixels. | Image Extraction | docx, jpeg | mcp |
| 5 | `apply-lossless-compression-to-tiff-images-extracted-from-rtf-files-and-store-them-in-an-ar.cs` | Apply lossless compression to TIFF images extracted from RTF files and store them in an archive. | Image Extraction | tiff | mcp |
| 6 | `insert-a-high-resolution-png-image-into-a-specific-paragraph-of-an-existing-docx-file.cs` | Insert a high-resolution PNG image into a specific paragraph of an existing DOCX file. | Image Insertion And Shape | docx, png | mcp |
| 7 | `replace-all-gif-images-in-a-word-document-with-equivalent-png-versions-using-a-custom-mapp.cs` | Replace all GIF images in a Word document with equivalent PNG versions using a custom mapping. | Image Conversion And Manipulation | docx, png, gif | mcp |
| 8 | `export-images-from-the-header-and-footer-sections-of-an-odt-document-into-separate-folders.cs` | Export images from the header and footer sections of an ODT document into separate folders. | Input Bootstrap | docx | mcp |
| 9 | `load-a-pdf-file-extract-embedded-images-and-convert-them-to-jpeg-with-85-quality.cs` | Load a PDF file, extract embedded images, and convert them to JPEG with 85% quality. | Image Extraction | pdf, jpeg | mcp |
| 10 | `create-a-memory-stream-containing-a-bmp-image-extracted-from-a-docx-and-pass-it-to-the-api.cs` | Create a memory stream containing a BMP image extracted from a DOCX and pass it to the API. | Image Extraction | docx, bmp | mcp |
| 11 | `generate-a-thumbnail-png-for-each-image-extracted-from-an-html-document-while-maintaining.cs` | Generate a thumbnail PNG for each image extracted from an HTML document while maintaining aspect ratio. | Image Extraction | html, png | mcp |
| 12 | `batch-convert-all-extracted-images-from-a-collection-of-word-files-to-webp-format-for-web.cs` | Batch convert all extracted images from a collection of Word files to WebP format for web use. | Image Extraction | webp | mcp |
| 13 | `apply-a-grayscale-filter-to-all-jpeg-images-extracted-from-doc-files-before-saving-them.cs` | Apply a grayscale filter to all JPEG images extracted from DOC files before saving them. | Image Extraction | doc, jpeg | mcp |
| 14 | `extract-images-from-tables-within-a-docx-document-and-store-them-in-a-zip-archive.cs` | Extract images from tables within a DOCX document and store them in a zip archive. | Image Extraction | docx | llm |
| 15 | `load-a-docm-file-extract-embedded-images-and-rename-each-file-using-its-original-shape-nam.cs` | Load a DOCM file, extract embedded images, and rename each file using its original shape name. | Image Extraction | docx | mcp |
| 16 | `convert-extracted-bmp-images-to-png-format-while-reducing-color-depth-to-256-colors.cs` | Convert extracted BMP images to PNG format while reducing color depth to 256 colors. | Image Extraction | png, bmp | mcp |
| 17 | `extract-images-from-footnotes-in-a-word-document-and-save-them-as-separate-jpeg-files.cs` | Extract images from footnotes in a Word document and save them as separate JPEG files. | Image Extraction | docx, jpeg | mcp |
| 18 | `batch-process-multiple-docx-files-extracting-images-and-generating-a-csv-report-of-image-m.cs` | Batch process multiple DOCX files, extracting images and generating a CSV report of image metadata. | Image Extraction | docx, csv | mcp |
| 19 | `resize-extracted-png-images-to-a-fixed-height-of-600-pixels-while-preserving-width-proport.cs` | Resize extracted PNG images to a fixed height of 600 pixels while preserving width proportionally. | Image Extraction | png | mcp |
| 20 | `apply-exif-orientation-correction-to-jpeg-images-extracted-from-scanned-pdf-documents-befo.cs` | Apply EXIF orientation correction to JPEG images extracted from scanned PDF documents before saving. | Image Extraction | pdf, jpeg | mcp |
| 21 | `load-a-docx-template-replace-placeholder-images-with-dynamically-generated-charts-and-save.cs` | Load a DOCX template, replace placeholder images with dynamically generated charts, and save the document. | General Image Workflow | docx | mcp |
| 22 | `extract-images-from-smartart-shapes-in-a-word-document-and-convert-them-to-svg-format.cs` | Extract images from SmartArt shapes in a Word document and convert them to SVG format. | Image Extraction | docx, svg | existing_repo |
| 23 | `batch-extract-images-from-a-set-of-odt-files-and-organize-them-by-original-document-name.cs` | Batch extract images from a set of ODT files and organize them by original document name. | Image Extraction | docx | mcp |
| 24 | `compress-extracted-jpeg-images-using-progressive-encoding-to-reduce-file-size-for-web-deli.cs` | Compress extracted JPEG images using progressive encoding to reduce file size for web delivery. | Image Extraction | jpeg | mcp |
| 25 | `load-a-docx-file-extract-images-and-embed-them-into-a-new-html-page-using-base64-data.cs` | Load a DOCX file, extract images, and embed them into a new HTML page using base64 data. | Image Extraction | docx, html | mcp |
| 26 | `convert-extracted-gif-animations-to-individual-png-frames-and-store-them-in-sequentially-n.cs` | Convert extracted GIF animations to individual PNG frames and store them in sequentially numbered files. | Image Extraction | png, gif | mcp |
| 27 | `apply-a-watermark-text-overlay-to-all-png-images-extracted-from-a-collection-of-word-docum.cs` | Apply a watermark text overlay to all PNG images extracted from a collection of Word documents. | Image Extraction | docx, png | mcp |
| 28 | `extract-images-from-comments-in-a-docx-file-and-save-them-using-the-comment-identifier-as.cs` | Extract images from comments in a DOCX file and save them using the comment identifier as filename. | Image Extraction | docx | mcp |
| 29 | `batch-convert-extracted-tiff-images-to-jpeg-with-90-quality-and-store-them-in-the-output-d.cs` | Batch convert extracted TIFF images to JPEG with 90% quality and store them in the output directory. | Image Extraction | jpeg, tiff | mcp |
| 30 | `load-a-doc-file-extract-all-embedded-vector-images-and-convert-them-to-emf-format.cs` | Load a DOC file, extract all embedded vector images, and convert them to EMF format. | Image Extraction | doc, emf | existing_repo |
| 31 | `resize-extracted-bmp-images-to-a-maximum-file-size-of-200-kilobytes-using-quality-scaling.cs` | Resize extracted BMP images to a maximum file size of 200 kilobytes using quality scaling. | Image Extraction | bmp | mcp |
| 32 | `extract-images-from-a-docx-document-s-custom-xml-parts-and-save-them-with-original-filenam.cs` | Extract images from a DOCX document's custom XML parts and save them with original filenames. | Image Extraction | docx, xml | mcp |
| 33 | `batch-process-docx-files-extracting-images-and-generating-a-json-manifest-containing-image.cs` | Batch process DOCX files, extracting images and generating a JSON manifest containing image dimensions. | Image Extraction | docx, json | mcp |
| 34 | `convert-extracted-png-images-to-grayscale-bmp-files-for-archival-purposes-in-a-secure-fold.cs` | Convert extracted PNG images to grayscale BMP files for archival purposes in a secure folder. | Image Extraction | png, bmp | mcp |
| 35 | `load-a-word-document-replace-all-low-resolution-images-with-high-resolution-versions-from.cs` | Load a Word document, replace all low-resolution images with high-resolution versions from a library. | General Image Workflow | docx | mcp |
| 36 | `extract-images-from-a-docx-file-s-background-settings-and-store-them-as-separate-png-files.cs` | Extract images from a DOCX file's background settings and store them as separate PNG files. | Image Extraction | docx, png | mcp |
| 37 | `batch-extract-images-from-a-set-of-pdf-files-and-rename-them-using-the-source-document-tit.cs` | Batch extract images from a set of PDF files and rename them using the source document title. | Image Extraction | pdf | mcp |
| 38 | `apply-a-blur-filter-to-all-extracted-jpeg-images-before-re-embedding-them-into-a-new-word.cs` | Apply a blur filter to all extracted JPEG images before re-embedding them into a new Word file. | Image Extraction | jpeg | mcp |
| 39 | `extract-images-from-a-docx-file-s-chart-objects-and-convert-them-to-high-resolution-png.cs` | Extract images from a DOCX file's chart objects and convert them to high-resolution PNG. | Image Extraction | docx, png | mcp |
| 40 | `resize-extracted-images-to-square-dimensions-of-500-by-500-pixels-adding-padding-where-nec.cs` | Resize extracted images to square dimensions of 500 by 500 pixels, adding padding where necessary. | Image Extraction | docx | mcp |
| 41 | `batch-convert-extracted-gif-images-to-animated-webp-files-while-preserving-original-frame.cs` | Batch convert extracted GIF images to animated WebP files while preserving original frame delays. | Image Extraction | gif, webp | mcp |
| 42 | `load-a-docx-file-extract-images-and-generate-an-excel-spreadsheet-listing-image-metadata.cs` | Load a DOCX file, extract images, and generate an Excel spreadsheet listing image metadata. | Image Extraction | docx | llm |
| 43 | `extract-images-from-a-word-document-s-equation-objects-and-save-them-as-png-for-external-u.cs` | Extract images from a Word document's equation objects and save them as PNG for external use. | Image Extraction | docx, png | llm |
| 44 | `apply-lossless-png-compression-to-all-extracted-images-and-compare-file-size-reduction-sta.cs` | Apply lossless PNG compression to all extracted images and compare file size reduction statistics. | Image Extraction | png | mcp |
| 45 | `batch-extract-images-from-doc-files-and-organize-them-into-subfolders-based-on-image-forma.cs` | Batch extract images from DOC files and organize them into subfolders based on image format type. | Image Extraction | doc | mcp |
| 46 | `convert-extracted-tiff-images-to-pdf-pages-each-image-occupying-a-full-page-in-the-output.cs` | Convert extracted TIFF images to PDF pages, each image occupying a full page in the output. | Image Extraction | pdf, tiff | mcp |
| 47 | `load-a-docx-document-replace-all-images-larger-than-2-mb-with-compressed-jpeg-equivalents.cs` | Load a DOCX document, replace all images larger than 2 MB with compressed JPEG equivalents. | Image Conversion And Manipulation | docx, jpeg | existing_repo |
| 48 | `extract-images-from-a-word-document-s-embedded-ole-objects-and-save-them-using-the-ole-ide.cs` | Extract images from a Word document's embedded OLE objects and save them using the OLE identifier. | Image Extraction | docx | mcp |
| 49 | `batch-process-a-collection-of-docx-files-extracting-images-and-creating-a-summary-pdf-cata.cs` | Batch process a collection of DOCX files, extracting images and creating a summary PDF catalog. | Image Extraction | docx, pdf | mcp |
| 50 | `resize-extracted-jpeg-images-to-1024x768-resolution-while-maintaining-original-aspect-rati.cs` | Resize extracted JPEG images to 1024x768 resolution while maintaining original aspect ratio. | Image Extraction | jpeg | mcp |
| 51 | `apply-a-sepia-tone-effect-to-all-png-images-extracted-from-a-set-of-word-documents.cs` | Apply a sepia tone effect to all PNG images extracted from a set of Word documents. | Image Extraction | docx, png | mcp |
| 52 | `extract-images-from-a-docx-file-s-custom-document-properties-and-store-them-with-property.cs` | Extract images from a DOCX file's custom document properties and store them with property names. | Image Extraction | docx | existing_repo |
| 53 | `batch-convert-extracted-bmp-images-to-jpeg-with-80-quality-and-log-conversion-results.cs` | Batch convert extracted BMP images to JPEG with 80% quality and log conversion results. | Image Extraction | jpeg, bmp | mcp |
| 54 | `load-a-word-document-extract-images-and-embed-them-into-a-powerpoint-presentation-automati.cs` | Load a Word document, extract images, and embed them into a PowerPoint presentation automatically. | Image Extraction | docx | llm |
| 55 | `extract-images-from-a-docx-file-s-content-controls-and-rename-files-using-control-identifi.cs` | Extract images from a DOCX file's content controls and rename files using control identifiers. | Image Extraction | docx | existing_repo |
| 56 | `resize-extracted-gif-images-to-a-maximum-width-of-300-pixels-while-preserving-animation-fr.cs` | Resize extracted GIF images to a maximum width of 300 pixels while preserving animation frames. | Image Extraction | gif | existing_repo |
| 57 | `apply-a-contrast-enhancement-filter-to-all-extracted-png-images-before-saving-them-to-disk.cs` | Apply a contrast enhancement filter to all extracted PNG images before saving them to disk. | Image Extraction | png | mcp |
| 58 | `batch-extract-images-from-a-set-of-odt-files-and-generate-a-markdown-gallery-with-thumbnai.cs` | Batch extract images from a set of ODT files and generate a markdown gallery with thumbnails. | Image Extraction | docx | llm |
| 59 | `convert-extracted-jpeg-images-to-grayscale-bmp-files-and-store-them-in-a-secure-archive.cs` | Convert extracted JPEG images to grayscale BMP files and store them in a secure archive. | Image Extraction | jpeg, bmp | mcp |
| 60 | `load-a-docx-template-programmatically-insert-images-from-a-database-and-save-the-result.cs` | Load a DOCX template, programmatically insert images from a database, and save the result. | Image Insertion And Shape | docx | mcp |
| 61 | `extract-images-from-a-word-document-s-footnote-references-and-name-files-using-footnote-nu.cs` | Extract images from a Word document's footnote references and name files using footnote numbers. | Image Extraction | docx | mcp |
| 62 | `batch-process-doc-files-extracting-images-and-creating-a-zip-archive-containing-all-extrac.cs` | Batch process DOC files, extracting images and creating a zip archive containing all extracted files. | Image Extraction | doc | mcp |
| 63 | `resize-extracted-png-images-to-50-of-original-dimensions-for-quick-preview-generation.cs` | Resize extracted PNG images to 50% of original dimensions for quick preview generation. | Image Extraction | png | mcp |
| 64 | `apply-a-vignette-effect-to-all-jpeg-images-extracted-from-a-collection-of-word-documents.cs` | Apply a vignette effect to all JPEG images extracted from a collection of Word documents. | Image Extraction | docx, jpeg | mcp |
| 65 | `extract-images-from-a-docx-file-s-embedded-video-thumbnails-and-save-them-as-png-files.cs` | Extract images from a DOCX file's embedded video thumbnails and save them as PNG files. | Image Extraction | docx, png | mcp |
| 66 | `batch-convert-extracted-tiff-images-to-lossless-png-while-preserving-metadata-such-as-dpi.cs` | Batch convert extracted TIFF images to lossless PNG while preserving metadata such as DPI. | Image Extraction | png, tiff | mcp |
| 67 | `load-a-word-document-extract-images-and-generate-an-html-gallery-page-with-lightbox-suppor.cs` | Load a Word document, extract images, and generate an HTML gallery page with lightbox support. | Image Extraction | docx, html | mcp |
| 68 | `extract-images-from-a-docx-file-s-custom-xml-schema-and-map-them-to-external-resource-ids.cs` | Extract images from a DOCX file's custom XML schema and map them to external resource IDs. | Image Extraction | docx, xml | existing_repo |
| 69 | `resize-extracted-bmp-images-to-a-fixed-1024-pixel-width-adjusting-height-proportionally.cs` | Resize extracted BMP images to a fixed 1024 pixel width, adjusting height proportionally. | Image Extraction | bmp | mcp |
| 70 | `apply-a-sharpening-filter-to-all-png-images-extracted-from-word-documents-before-re-saving.cs` | Apply a sharpening filter to all PNG images extracted from Word documents before re-saving them. | Image Extraction | docx, png | llm |
| 71 | `batch-extract-images-from-a-set-of-docx-files-and-create-a-pdf-catalog-with-thumbnails.cs` | Batch extract images from a set of DOCX files and create a PDF catalog with thumbnails. | Image Extraction | docx, pdf | mcp |
| 72 | `convert-extracted-jpeg-images-to-high-resolution-tiff-for-archival-storage-with-lzw-compre.cs` | Convert extracted JPEG images to high-resolution TIFF for archival storage with LZW compression. | Image Extraction | jpeg, tiff | mcp |
| 73 | `load-a-docx-document-replace-all-images-with-placeholders-and-export-the-modified-document.cs` | Load a DOCX document, replace all images with placeholders, and export the modified document. | General Image Workflow | docx | mcp |
| 74 | `extract-images-from-a-word-document-s-section-breaks-and-save-them-using-section-identifie.cs` | Extract images from a Word document's section breaks and save them using section identifiers. | Image Extraction | docx | mcp |
| 75 | `batch-process-multiple-odt-files-extracting-images-and-generating-a-consolidated-json-mani.cs` | Batch process multiple ODT files, extracting images and generating a consolidated JSON manifest. | Image Extraction | json | mcp |
| 76 | `resize-extracted-gif-images-to-200x200-pixels-and-convert-them-to-static-png-format.cs` | Resize extracted GIF images to 200x200 pixels and convert them to static PNG format. | Image Extraction | png, gif | mcp |
| 77 | `apply-a-color-inversion-filter-to-all-extracted-png-images-for-visual-effect-testing.cs` | Apply a color inversion filter to all extracted PNG images for visual effect testing. | Image Extraction | png | mcp |
| 78 | `extract-images-from-a-docx-file-s-embedded-audio-cover-art-and-store-them-as-jpeg-files.cs` | Extract images from a DOCX file's embedded audio cover art and store them as JPEG files. | Image Extraction | docx, jpeg | mcp |
| 79 | `batch-convert-extracted-bmp-images-to-webp-with-lossless-compression-and-log-conversion-de.cs` | Batch convert extracted BMP images to WebP with lossless compression and log conversion details. | Image Extraction | bmp, webp | llm |
| 80 | `load-a-word-document-extract-images-and-create-a-powershell-script-to-re-embed-them-later.cs` | Load a Word document, extract images, and create a PowerShell script to re-embed them later. | Image Extraction | docx | mcp |
| 81 | `extract-images-from-a-docx-file-s-embedded-map-objects-and-save-them-as-high-resolution-pn.cs` | Extract images from a DOCX file's embedded map objects and save them as high-resolution PNG. | Image Extraction | docx, png | mcp |
| 82 | `resize-extracted-jpeg-images-to-a-maximum-file-size-of-500kb-using-adaptive-quality.cs` | Resize extracted JPEG images to a maximum file size of 500KB using adaptive quality. | Image Extraction | jpeg | mcp |
| 83 | `apply-a-border-of-5-pixels-in-red-color-to-all-extracted-png-images-before-saving.cs` | Apply a border of 5 pixels in red color to all extracted PNG images before saving. | Image Extraction | png | mcp |
| 84 | `batch-extract-images-from-a-collection-of-doc-files-and-generate-an-excel-index-with-file.cs` | Batch extract images from a collection of DOC files and generate an Excel index with file paths. | Image Extraction | doc | mcp |
| 85 | `convert-extracted-tiff-images-to-grayscale-jpeg-for-use-in-low-bandwidth-environments.cs` | Convert extracted TIFF images to grayscale JPEG for use in low-bandwidth environments. | Image Extraction | jpeg, tiff | mcp |
| 86 | `load-a-docx-template-programmatically-insert-images-from-a-rest-api-and-save-as-pdf.cs` | Load a DOCX template, programmatically insert images from a REST API, and save as PDF. | Image Insertion And Shape | docx, pdf | mcp |
| 87 | `extract-images-from-a-word-document-s-embedded-chart-data-and-export-them-as-svg-files.cs` | Extract images from a Word document's embedded chart data and export them as SVG files. | Image Extraction | docx, svg | mcp |
| 88 | `batch-process-docx-files-extracting-images-and-creating-a-summary-csv-containing-image-siz.cs` | Batch process DOCX files, extracting images and creating a summary CSV containing image sizes. | Image Extraction | docx, csv | existing_repo |
| 89 | `resize-extracted-png-images-to-75-of-original-dimensions-for-quick-preview-generation.cs` | Resize extracted PNG images to 75% of original dimensions for quick preview generation. | Image Extraction | png | existing_repo |
| 90 | `apply-a-gaussian-blur-to-all-extracted-jpeg-images-before-embedding-them-into-a-new-docume.cs` | Apply a Gaussian blur to all extracted JPEG images before embedding them into a new document. | Image Extraction | jpeg | existing_repo |
| 91 | `extract-images-from-a-docx-file-s-embedded-3d-model-thumbnails-and-save-as-png.cs` | Extract images from a DOCX file's embedded 3D model thumbnails and save as PNG. | Image Extraction | docx, png | existing_repo |
| 92 | `batch-convert-extracted-gif-animations-to-mp4-video-clips-while-preserving-frame-order.cs` | Batch convert extracted GIF animations to MP4 video clips while preserving frame order. | Image Extraction | gif | existing_repo |
| 93 | `load-a-word-document-extract-images-and-generate-a-latex-file-referencing-each-image.cs` | Load a Word document, extract images, and generate a LaTeX file referencing each image. | Image Extraction | docx | existing_repo |
| 94 | `extract-images-from-a-docx-file-s-custom-xml-data-and-map-them-to-database-records.cs` | Extract images from a DOCX file's custom XML data and map them to database records. | Image Extraction | docx, xml | existing_repo |
| 95 | `resize-extracted-bmp-images-to-640x480-resolution-for-use-in-legacy-applications.cs` | Resize extracted BMP images to 640x480 resolution for use in legacy applications. | Image Extraction | bmp | existing_repo |
| 96 | `apply-a-hue-rotation-of-180-degrees-to-all-extracted-png-images-for-color-testing.cs` | Apply a hue rotation of 180 degrees to all extracted PNG images for color testing. | Image Extraction | png | existing_repo |
| 97 | `batch-extract-images-from-a-set-of-odt-files-and-create-a-searchable-pdf-catalog.cs` | Batch extract images from a set of ODT files and create a searchable PDF catalog. | Image Extraction | pdf | existing_repo |
| 98 | `convert-extracted-jpeg-images-to-high-quality-webp-for-optimized-web-delivery.cs` | Convert extracted JPEG images to high-quality WebP for optimized web delivery. | Image Extraction | jpeg, webp | existing_repo |
| 99 | `load-a-docx-file-replace-all-images-with-placeholders-and-export-to-html-preserving-layout.cs` | Load a DOCX file, replace all images with placeholders, and export to HTML preserving layout. | General Image Workflow | docx, html | existing_repo |
| 100 | `extract-images-from-a-word-document-s-embedded-audio-tracks-and-save-them-as-jpeg-thumbnai.cs` | Extract images from a Word document's embedded audio tracks and save them as JPEG thumbnails. | Image Extraction | docx, jpeg | existing_repo |
| 101 | `batch-process-multiple-doc-files-extracting-images-and-generating-a-consolidated-pdf-repor.cs` | Batch process multiple DOC files, extracting images and generating a consolidated PDF report. | Image Extraction | doc, pdf | existing_repo |
| 102 | `resize-extracted-png-images-to-a-maximum-dimension-of-1200-pixels-maintaining-aspect-ratio.cs` | Resize extracted PNG images to a maximum dimension of 1200 pixels, maintaining aspect ratio. | Image Extraction | png | existing_repo |
| 103 | `apply-a-motion-blur-effect-to-all-extracted-jpeg-images-before-re-embedding-them.cs` | Apply a motion blur effect to all extracted JPEG images before re-embedding them. | Image Extraction | jpeg | existing_repo |
| 104 | `extract-images-from-a-docx-file-s-embedded-spreadsheet-charts-and-save-them-as-png-files.cs` | Extract images from a DOCX file's embedded spreadsheet charts and save them as PNG files. | Image Extraction | docx, png | existing_repo |
| 105 | `batch-convert-extracted-bmp-images-to-lossless-png-and-generate-a-checksum-file-for-verifi.cs` | Batch convert extracted BMP images to lossless PNG and generate a checksum file for verification. | Image Extraction | png, bmp | existing_repo |
| 106 | `load-a-word-document-extract-images-and-create-a-markdown-file-with-image-links.cs` | Load a Word document, extract images, and create a Markdown file with image links. | Image Extraction | docx | existing_repo |
| 107 | `extract-images-from-a-docx-file-s-embedded-video-frames-and-store-them-as-high-resolution.cs` | Extract images from a DOCX file's embedded video frames and store them as high-resolution PNG. | Image Extraction | docx, png | existing_repo |
| 108 | `resize-extracted-jpeg-images-to-50-width-and-height-for-thumbnail-generation-in-galleries.cs` | Resize extracted JPEG images to 50% width and height for thumbnail generation in galleries. | Image Extraction | jpeg | existing_repo |
| 109 | `apply-a-color-balance-adjustment-to-all-extracted-png-images-before-saving-them-to-the-out.cs` | Apply a color balance adjustment to all extracted PNG images before saving them to the output folder. | Image Extraction | png | existing_repo |
| 110 | `batch-extract-images-from-a-collection-of-docx-files-and-generate-an-html-index-page.cs` | Batch extract images from a collection of DOCX files and generate an HTML index page. | Image Extraction | docx, html | existing_repo |
| 111 | `convert-extracted-tiff-images-to-pdf-with-each-image-on-a-separate-page-and-embed-metadata.cs` | Convert extracted TIFF images to PDF with each image on a separate page and embed metadata. | Image Extraction | pdf, tiff | existing_repo |
| 112 | `load-a-docx-template-programmatically-insert-images-from-a-file-system-and-save-as-docx.cs` | Load a DOCX template, programmatically insert images from a file system, and save as DOCX. | Image Insertion And Shape | docx, doc | existing_repo |
| 113 | `extract-images-from-a-word-document-s-embedded-qr-codes-and-decode-them-for-data-extractio.cs` | Extract images from a Word document's embedded QR codes and decode them for data extraction. | Image Extraction | docx | existing_repo |
| 114 | `batch-process-doc-files-extracting-images-and-creating-a-compressed-zip-archive-with-passw.cs` | Batch process DOC files, extracting images and creating a compressed ZIP archive with password protection. | Image Extraction | doc | existing_repo |
| 115 | `resize-extracted-png-images-to-300x300-pixels-and-add-a-watermark-text-overlay.cs` | Resize extracted PNG images to 300x300 pixels and add a watermark text overlay. | Image Extraction | png | existing_repo |
| 116 | `apply-a-saturation-increase-of-20-to-all-extracted-jpeg-images-before-re-saving-them.cs` | Apply a saturation increase of 20% to all extracted JPEG images before re-saving them. | Image Extraction | jpeg | existing_repo |
| 117 | `extract-images-from-a-docx-file-s-embedded-audio-waveform-visualizations-and-save-as-png.cs` | Extract images from a DOCX file's embedded audio waveform visualizations and save as PNG. | Image Extraction | docx, png | existing_repo |
| 118 | `batch-convert-extracted-gif-images-to-animated-png-while-preserving-original-frame-timing.cs` | Batch convert extracted GIF images to animated PNG while preserving original frame timing. | Image Extraction | png, gif | existing_repo |
| 119 | `load-a-word-document-extract-images-and-generate-a-json-file-containing-base64-encoded-dat.cs` | Load a Word document, extract images, and generate a JSON file containing base64 encoded data. | Image Extraction | docx, json | existing_repo |
| 120 | `extract-images-from-a-docx-file-s-embedded-map-tiles-and-store-them-using-tile-coordinates.cs` | Extract images from a DOCX file's embedded map tiles and store them using tile coordinates. | Image Extraction | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Missing local image bootstrap

- Symptom: Examples assume an existing image file or source document already exists.
- Fix: Create deterministic local sample images and source documents before insertion, extraction, or conversion.

### Shape insertion at invalid location

- Symptom: Runtime failures such as invalid insertion location or self-append mistakes.
- Fix: Append image shapes only once and only into valid containers such as Paragraph.

### Unsafe image extraction

- Symptom: Code accesses ImageData on nodes that are not image-bearing shapes.
- Fix: Enumerate Shape nodes and require shape.HasImage before accessing or saving ImageData.

### Stream reuse bugs

- Symptom: Image or document stream operations fail because MemoryStream.Position was not reset.
- Fix: Always reset stream.Position = 0 before reuse.

### Context-specific image container mistakes

- Symptom: Comment, content-control, footnote, or property-based image tasks fail due to wrong document structure assumptions.
- Fix: Traverse the actual container nodes first, then shapes inside them, and insert shapes only through valid child containers.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.5.0 run.
- Preserve file-to-task traceability when updating this folder.
- Keep examples standalone and bootstrap local inputs inside the example whenever external sources are mentioned.
