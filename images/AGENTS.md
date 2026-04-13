---
name: images
description: Verified C# examples for image scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Images

## Purpose

This folder is a live, curated example set for image scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free image creation, insertion, extraction, conversion, and container-aware image handling using direct Aspose.Words APIs and deterministic local image assets.

## Non-negotiable conventions

- Use Aspose.Words Shape and ImageData APIs for image operations.
- Create local deterministic image inputs whenever the task implies an existing image, file, stream, folder, or batch source.
- Use DocumentBuilder.InsertImage for standard insertion tasks and Shape.ImageData for shape-based scenarios.
- Always check shape.HasImage before accessing ImageData.
- Do not use System.Drawing.
- Reset MemoryStream.Position before reuse.
- Validate that expected output files are actually written.

## Recommended workflow selection

- **Image extraction workflow**: 109 examples
- **Image insertion / shape workflow**: 4 examples
- **Image conversion / manipulation workflow**: 2 examples
- **Input-bootstrap workflow**: 1 examples
- **General image workflow**: 4 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. Source images, documents, streams, and folders must be bootstrapped locally whenever the task implies an existing input.
3. Image operations must use documented Aspose.Words APIs and valid document or container structures.
4. Requested output documents, extracted images, or converted image files must actually be written.
5. Container-specific image tasks should traverse the actual container nodes first, then image-bearing shapes inside them.

## File-to-task reference

- `load-a-docx-file-and-extract-all-embedded-images-into-a-specified-output-folder.cs`
  - Task: Load a DOCX file and extract all embedded images into a specified output folder.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `convert-each-extracted-png-image-to-jpeg-format-while-preserving-its-original-dimensions.cs`
  - Task: Convert each extracted PNG image to JPEG format while preserving its original dimensions.
  - Workflow: image-extraction
  - Outputs: png, jpeg
  - Selected engine: verified
- `batch-process-a-folder-of-doc-files-extracting-images-and-saving-them-as-bmp-files.cs`
  - Task: Batch process a folder of DOC files, extracting images and saving them as BMP files.
  - Workflow: image-extraction
  - Outputs: doc, bmp
  - Selected engine: verified
- `resize-all-jpeg-images-extracted-from-a-word-document-to-a-maximum-width-of-800-pixels.cs`
  - Task: Resize all JPEG images extracted from a Word document to a maximum width of 800 pixels.
  - Workflow: image-extraction
  - Outputs: docx, jpeg
  - Selected engine: verified
- `apply-lossless-compression-to-tiff-images-extracted-from-rtf-files-and-store-them-in-an-ar.cs`
  - Task: Apply lossless compression to TIFF images extracted from RTF files and store them in an archive.
  - Workflow: image-extraction
  - Outputs: tiff
  - Selected engine: verified
- `insert-a-high-resolution-png-image-into-a-specific-paragraph-of-an-existing-docx-file.cs`
  - Task: Insert a high‑resolution PNG image into a specific paragraph of an existing DOCX file.
  - Workflow: image-insertion-and-shape
  - Outputs: docx, png
  - Selected engine: verified
- `replace-all-gif-images-in-a-word-document-with-equivalent-png-versions-using-a-custom-mapp.cs`
  - Task: Replace all GIF images in a Word document with equivalent PNG versions using a custom mapping.
  - Workflow: image-conversion-and-manipulation
  - Outputs: docx, png, gif
  - Selected engine: verified
- `export-images-from-the-header-and-footer-sections-of-an-odt-document-into-separate-folders.cs`
  - Task: Export images from the header and footer sections of an ODT document into separate folders.
  - Workflow: input-bootstrap
  - Outputs: docx
  - Selected engine: verified
- `load-a-pdf-file-extract-embedded-images-and-convert-them-to-jpeg-with-85-quality.cs`
  - Task: Load a PDF file, extract embedded images, and convert them to JPEG with 85% quality.
  - Workflow: image-extraction
  - Outputs: pdf, jpeg
  - Selected engine: verified
- `create-a-memory-stream-containing-a-bmp-image-extracted-from-a-docx-and-pass-it-to-the-api.cs`
  - Task: Create a memory stream containing a BMP image extracted from a DOCX and pass it to the API.
  - Workflow: image-extraction
  - Outputs: docx, bmp
  - Selected engine: verified
- `generate-a-thumbnail-png-for-each-image-extracted-from-an-html-document-while-maintaining.cs`
  - Task: Generate a thumbnail PNG for each image extracted from an HTML document while maintaining aspect ratio.
  - Workflow: image-extraction
  - Outputs: html, png
  - Selected engine: verified
- `batch-convert-all-extracted-images-from-a-collection-of-word-files-to-webp-format-for-web.cs`
  - Task: Batch convert all extracted images from a collection of Word files to WebP format for web use.
  - Workflow: image-extraction
  - Outputs: webp
  - Selected engine: verified
- `apply-a-grayscale-filter-to-all-jpeg-images-extracted-from-doc-files-before-saving-them.cs`
  - Task: Apply a grayscale filter to all JPEG images extracted from DOC files before saving them.
  - Workflow: image-extraction
  - Outputs: doc, jpeg
  - Selected engine: verified
- `extract-images-from-tables-within-a-docx-document-and-store-them-in-a-zip-archive.cs`
  - Task: Extract images from tables within a DOCX document and store them in a zip archive.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `load-a-docm-file-extract-embedded-images-and-rename-each-file-using-its-original-shape-nam.cs`
  - Task: Load a DOCM file, extract embedded images, and rename each file using its original shape name.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `convert-extracted-bmp-images-to-png-format-while-reducing-color-depth-to-256-colors.cs`
  - Task: Convert extracted BMP images to PNG format while reducing color depth to 256 colors.
  - Workflow: image-extraction
  - Outputs: png, bmp
  - Selected engine: verified
- `extract-images-from-footnotes-in-a-word-document-and-save-them-as-separate-jpeg-files.cs`
  - Task: Extract images from footnotes in a Word document and save them as separate JPEG files.
  - Workflow: image-extraction
  - Outputs: docx, jpeg
  - Selected engine: verified
- `batch-process-multiple-docx-files-extracting-images-and-generating-a-csv-report-of-image-m.cs`
  - Task: Batch process multiple DOCX files, extracting images and generating a CSV report of image metadata.
  - Workflow: image-extraction
  - Outputs: docx, csv
  - Selected engine: verified
- `resize-extracted-png-images-to-a-fixed-height-of-600-pixels-while-preserving-width-proport.cs`
  - Task: Resize extracted PNG images to a fixed height of 600 pixels while preserving width proportionally.
  - Workflow: image-extraction
  - Outputs: png
  - Selected engine: verified
- `apply-exif-orientation-correction-to-jpeg-images-extracted-from-scanned-pdf-documents-befo.cs`
  - Task: Apply EXIF orientation correction to JPEG images extracted from scanned PDF documents before saving.
  - Workflow: image-extraction
  - Outputs: pdf, jpeg
  - Selected engine: verified
- `load-a-docx-template-replace-placeholder-images-with-dynamically-generated-charts-and-save.cs`
  - Task: Load a DOCX template, replace placeholder images with dynamically generated charts, and save the document.
  - Workflow: general-image-workflow
  - Outputs: docx
  - Selected engine: verified
- `extract-images-from-smartart-shapes-in-a-word-document-and-convert-them-to-svg-format.cs`
  - Task: Extract images from SmartArt shapes in a Word document and convert them to SVG format.
  - Workflow: image-extraction
  - Outputs: docx, svg
  - Selected engine: verified
- `batch-extract-images-from-a-set-of-odt-files-and-organize-them-by-original-document-name.cs`
  - Task: Batch extract images from a set of ODT files and organize them by original document name.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `compress-extracted-jpeg-images-using-progressive-encoding-to-reduce-file-size-for-web-deli.cs`
  - Task: Compress extracted JPEG images using progressive encoding to reduce file size for web delivery.
  - Workflow: image-extraction
  - Outputs: jpeg
  - Selected engine: verified
- `load-a-docx-file-extract-images-and-embed-them-into-a-new-html-page-using-base64-data.cs`
  - Task: Load a DOCX file, extract images, and embed them into a new HTML page using base64 data.
  - Workflow: image-extraction
  - Outputs: docx, html
  - Selected engine: verified
- `convert-extracted-gif-animations-to-individual-png-frames-and-store-them-in-sequentially-n.cs`
  - Task: Convert extracted GIF animations to individual PNG frames and store them in sequentially numbered files.
  - Workflow: image-extraction
  - Outputs: png, gif
  - Selected engine: verified
- `apply-a-watermark-text-overlay-to-all-png-images-extracted-from-a-collection-of-word-docum.cs`
  - Task: Apply a watermark text overlay to all PNG images extracted from a collection of Word documents.
  - Workflow: image-extraction
  - Outputs: docx, png
  - Selected engine: verified
- `extract-images-from-comments-in-a-docx-file-and-save-them-using-the-comment-identifier-as.cs`
  - Task: Extract images from comments in a DOCX file and save them using the comment identifier as filename.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `batch-convert-extracted-tiff-images-to-jpeg-with-90-quality-and-store-them-in-the-output-d.cs`
  - Task: Batch convert extracted TIFF images to JPEG with 90% quality and store them in the output directory.
  - Workflow: image-extraction
  - Outputs: jpeg, tiff
  - Selected engine: verified
- `load-a-doc-file-extract-all-embedded-vector-images-and-convert-them-to-emf-format.cs`
  - Task: Load a DOC file, extract all embedded vector images, and convert them to EMF format.
  - Workflow: image-extraction
  - Outputs: doc, emf
  - Selected engine: verified
- `resize-extracted-bmp-images-to-a-maximum-file-size-of-200-kilobytes-using-quality-scaling.cs`
  - Task: Resize extracted BMP images to a maximum file size of 200 kilobytes using quality scaling.
  - Workflow: image-extraction
  - Outputs: bmp
  - Selected engine: verified
- `extract-images-from-a-docx-document-s-custom-xml-parts-and-save-them-with-original-filenam.cs`
  - Task: Extract images from a DOCX document's custom XML parts and save them with original filenames.
  - Workflow: image-extraction
  - Outputs: docx, xml
  - Selected engine: verified
- `batch-process-docx-files-extracting-images-and-generating-a-json-manifest-containing-image.cs`
  - Task: Batch process DOCX files, extracting images and generating a JSON manifest containing image dimensions.
  - Workflow: image-extraction
  - Outputs: docx, json
  - Selected engine: verified
- `convert-extracted-png-images-to-grayscale-bmp-files-for-archival-purposes-in-a-secure-fold.cs`
  - Task: Convert extracted PNG images to grayscale BMP files for archival purposes in a secure folder.
  - Workflow: image-extraction
  - Outputs: png, bmp
  - Selected engine: verified
- `load-a-word-document-replace-all-low-resolution-images-with-high-resolution-versions-from.cs`
  - Task: Load a Word document, replace all low‑resolution images with high‑resolution versions from a library.
  - Workflow: general-image-workflow
  - Outputs: docx
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-background-settings-and-store-them-as-separate-png-files.cs`
  - Task: Extract images from a DOCX file's background settings and store them as separate PNG files.
  - Workflow: image-extraction
  - Outputs: docx, png
  - Selected engine: verified
- `batch-extract-images-from-a-set-of-pdf-files-and-rename-them-using-the-source-document-tit.cs`
  - Task: Batch extract images from a set of PDF files and rename them using the source document title.
  - Workflow: image-extraction
  - Outputs: pdf
  - Selected engine: verified
- `apply-a-blur-filter-to-all-extracted-jpeg-images-before-re-embedding-them-into-a-new-word.cs`
  - Task: Apply a blur filter to all extracted JPEG images before re‑embedding them into a new Word file.
  - Workflow: image-extraction
  - Outputs: jpeg
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-chart-objects-and-convert-them-to-high-resolution-png.cs`
  - Task: Extract images from a DOCX file's chart objects and convert them to high‑resolution PNG.
  - Workflow: image-extraction
  - Outputs: docx, png
  - Selected engine: verified
- `resize-extracted-images-to-square-dimensions-of-500-by-500-pixels-adding-padding-where-nec.cs`
  - Task: Resize extracted images to square dimensions of 500 by 500 pixels, adding padding where necessary.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `batch-convert-extracted-gif-images-to-animated-webp-files-while-preserving-original-frame.cs`
  - Task: Batch convert extracted GIF images to animated WebP files while preserving original frame delays.
  - Workflow: image-extraction
  - Outputs: gif, webp
  - Selected engine: verified
- `load-a-docx-file-extract-images-and-generate-an-excel-spreadsheet-listing-image-metadata.cs`
  - Task: Load a DOCX file, extract images, and generate an Excel spreadsheet listing image metadata.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-images-from-a-word-document-s-equation-objects-and-save-them-as-png-for-external-u.cs`
  - Task: Extract images from a Word document's equation objects and save them as PNG for external use.
  - Workflow: image-extraction
  - Outputs: docx, png
  - Selected engine: verified
- `apply-lossless-png-compression-to-all-extracted-images-and-compare-file-size-reduction-sta.cs`
  - Task: Apply lossless PNG compression to all extracted images and compare file size reduction statistics.
  - Workflow: image-extraction
  - Outputs: png
  - Selected engine: verified
- `batch-extract-images-from-doc-files-and-organize-them-into-subfolders-based-on-image-forma.cs`
  - Task: Batch extract images from DOC files and organize them into subfolders based on image format type.
  - Workflow: image-extraction
  - Outputs: doc
  - Selected engine: verified
- `convert-extracted-tiff-images-to-pdf-pages-each-image-occupying-a-full-page-in-the-output.cs`
  - Task: Convert extracted TIFF images to PDF pages, each image occupying a full page in the output.
  - Workflow: image-extraction
  - Outputs: pdf, tiff
  - Selected engine: verified
- `load-a-docx-document-replace-all-images-larger-than-2-mb-with-compressed-jpeg-equivalents.cs`
  - Task: Load a DOCX document, replace all images larger than 2 MB with compressed JPEG equivalents.
  - Workflow: image-conversion-and-manipulation
  - Outputs: docx, jpeg
  - Selected engine: verified
- `extract-images-from-a-word-document-s-embedded-ole-objects-and-save-them-using-the-ole-ide.cs`
  - Task: Extract images from a Word document's embedded OLE objects and save them using the OLE identifier.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `batch-process-a-collection-of-docx-files-extracting-images-and-creating-a-summary-pdf-cata.cs`
  - Task: Batch process a collection of DOCX files, extracting images and creating a summary PDF catalog.
  - Workflow: image-extraction
  - Outputs: docx, pdf
  - Selected engine: verified
- `resize-extracted-jpeg-images-to-1024x768-resolution-while-maintaining-original-aspect-rati.cs`
  - Task: Resize extracted JPEG images to 1024x768 resolution while maintaining original aspect ratio.
  - Workflow: image-extraction
  - Outputs: jpeg
  - Selected engine: verified
- `apply-a-sepia-tone-effect-to-all-png-images-extracted-from-a-set-of-word-documents.cs`
  - Task: Apply a sepia tone effect to all PNG images extracted from a set of Word documents.
  - Workflow: image-extraction
  - Outputs: docx, png
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-custom-document-properties-and-store-them-with-property.cs`
  - Task: Extract images from a DOCX file's custom document properties and store them with property names.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `batch-convert-extracted-bmp-images-to-jpeg-with-80-quality-and-log-conversion-results.cs`
  - Task: Batch convert extracted BMP images to JPEG with 80% quality and log conversion results.
  - Workflow: image-extraction
  - Outputs: jpeg, bmp
  - Selected engine: verified
- `load-a-word-document-extract-images-and-embed-them-into-a-powerpoint-presentation-automati.cs`
  - Task: Load a Word document, extract images, and embed them into a PowerPoint presentation automatically.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-content-controls-and-rename-files-using-control-identifi.cs`
  - Task: Extract images from a DOCX file's content controls and rename files using control identifiers.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `resize-extracted-gif-images-to-a-maximum-width-of-300-pixels-while-preserving-animation-fr.cs`
  - Task: Resize extracted GIF images to a maximum width of 300 pixels while preserving animation frames.
  - Workflow: image-extraction
  - Outputs: gif
  - Selected engine: verified
- `apply-a-contrast-enhancement-filter-to-all-extracted-png-images-before-saving-them-to-disk.cs`
  - Task: Apply a contrast enhancement filter to all extracted PNG images before saving them to disk.
  - Workflow: image-extraction
  - Outputs: png
  - Selected engine: verified
- `batch-extract-images-from-a-set-of-odt-files-and-generate-a-markdown-gallery-with-thumbnai.cs`
  - Task: Batch extract images from a set of ODT files and generate a markdown gallery with thumbnails.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `convert-extracted-jpeg-images-to-grayscale-bmp-files-and-store-them-in-a-secure-archive.cs`
  - Task: Convert extracted JPEG images to grayscale BMP files and store them in a secure archive.
  - Workflow: image-extraction
  - Outputs: jpeg, bmp
  - Selected engine: verified
- `load-a-docx-template-programmatically-insert-images-from-a-database-and-save-the-result.cs`
  - Task: Load a DOCX template, programmatically insert images from a database, and save the result.
  - Workflow: image-insertion-and-shape
  - Outputs: docx
  - Selected engine: verified
- `extract-images-from-a-word-document-s-footnote-references-and-name-files-using-footnote-nu.cs`
  - Task: Extract images from a Word document's footnote references and name files using footnote numbers.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `batch-process-doc-files-extracting-images-and-creating-a-zip-archive-containing-all-extrac.cs`
  - Task: Batch process DOC files, extracting images and creating a zip archive containing all extracted files.
  - Workflow: image-extraction
  - Outputs: doc
  - Selected engine: verified
- `resize-extracted-png-images-to-50-of-original-dimensions-for-quick-preview-generation.cs`
  - Task: Resize extracted PNG images to 50% of original dimensions for quick preview generation.
  - Workflow: image-extraction
  - Outputs: png
  - Selected engine: verified
- `apply-a-vignette-effect-to-all-jpeg-images-extracted-from-a-collection-of-word-documents.cs`
  - Task: Apply a vignette effect to all JPEG images extracted from a collection of Word documents.
  - Workflow: image-extraction
  - Outputs: docx, jpeg
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-embedded-video-thumbnails-and-save-them-as-png-files.cs`
  - Task: Extract images from a DOCX file's embedded video thumbnails and save them as PNG files.
  - Workflow: image-extraction
  - Outputs: docx, png
  - Selected engine: verified
- `batch-convert-extracted-tiff-images-to-lossless-png-while-preserving-metadata-such-as-dpi.cs`
  - Task: Batch convert extracted TIFF images to lossless PNG while preserving metadata such as DPI.
  - Workflow: image-extraction
  - Outputs: png, tiff
  - Selected engine: verified
- `load-a-word-document-extract-images-and-generate-an-html-gallery-page-with-lightbox-suppor.cs`
  - Task: Load a Word document, extract images, and generate an HTML gallery page with lightbox support.
  - Workflow: image-extraction
  - Outputs: docx, html
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-custom-xml-schema-and-map-them-to-external-resource-ids.cs`
  - Task: Extract images from a DOCX file's custom XML schema and map them to external resource IDs.
  - Workflow: image-extraction
  - Outputs: docx, xml
  - Selected engine: verified
- `resize-extracted-bmp-images-to-a-fixed-1024-pixel-width-adjusting-height-proportionally.cs`
  - Task: Resize extracted BMP images to a fixed 1024 pixel width, adjusting height proportionally.
  - Workflow: image-extraction
  - Outputs: bmp
  - Selected engine: verified
- `apply-a-sharpening-filter-to-all-png-images-extracted-from-word-documents-before-re-saving.cs`
  - Task: Apply a sharpening filter to all PNG images extracted from Word documents before re‑saving them.
  - Workflow: image-extraction
  - Outputs: docx, png
  - Selected engine: verified
- `batch-extract-images-from-a-set-of-docx-files-and-create-a-pdf-catalog-with-thumbnails.cs`
  - Task: Batch extract images from a set of DOCX files and create a PDF catalog with thumbnails.
  - Workflow: image-extraction
  - Outputs: docx, pdf
  - Selected engine: verified
- `convert-extracted-jpeg-images-to-high-resolution-tiff-for-archival-storage-with-lzw-compre.cs`
  - Task: Convert extracted JPEG images to high‑resolution TIFF for archival storage with LZW compression.
  - Workflow: image-extraction
  - Outputs: jpeg, tiff
  - Selected engine: verified
- `load-a-docx-document-replace-all-images-with-placeholders-and-export-the-modified-document.cs`
  - Task: Load a DOCX document, replace all images with placeholders, and export the modified document.
  - Workflow: general-image-workflow
  - Outputs: docx
  - Selected engine: verified
- `extract-images-from-a-word-document-s-section-breaks-and-save-them-using-section-identifie.cs`
  - Task: Extract images from a Word document's section breaks and save them using section identifiers.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `batch-process-multiple-odt-files-extracting-images-and-generating-a-consolidated-json-mani.cs`
  - Task: Batch process multiple ODT files, extracting images and generating a consolidated JSON manifest.
  - Workflow: image-extraction
  - Outputs: json
  - Selected engine: verified
- `resize-extracted-gif-images-to-200x200-pixels-and-convert-them-to-static-png-format.cs`
  - Task: Resize extracted GIF images to 200x200 pixels and convert them to static PNG format.
  - Workflow: image-extraction
  - Outputs: png, gif
  - Selected engine: verified
- `apply-a-color-inversion-filter-to-all-extracted-png-images-for-visual-effect-testing.cs`
  - Task: Apply a color inversion filter to all extracted PNG images for visual effect testing.
  - Workflow: image-extraction
  - Outputs: png
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-embedded-audio-cover-art-and-store-them-as-jpeg-files.cs`
  - Task: Extract images from a DOCX file's embedded audio cover art and store them as JPEG files.
  - Workflow: image-extraction
  - Outputs: docx, jpeg
  - Selected engine: verified
- `batch-convert-extracted-bmp-images-to-webp-with-lossless-compression-and-log-conversion-de.cs`
  - Task: Batch convert extracted BMP images to WebP with lossless compression and log conversion details.
  - Workflow: image-extraction
  - Outputs: bmp, webp
  - Selected engine: verified
- `load-a-word-document-extract-images-and-create-a-powershell-script-to-re-embed-them-later.cs`
  - Task: Load a Word document, extract images, and create a PowerShell script to re‑embed them later.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-embedded-map-objects-and-save-them-as-high-resolution-pn.cs`
  - Task: Extract images from a DOCX file's embedded map objects and save them as high‑resolution PNG.
  - Workflow: image-extraction
  - Outputs: docx, png
  - Selected engine: verified
- `resize-extracted-jpeg-images-to-a-maximum-file-size-of-500kb-using-adaptive-quality.cs`
  - Task: Resize extracted JPEG images to a maximum file size of 500KB using adaptive quality.
  - Workflow: image-extraction
  - Outputs: jpeg
  - Selected engine: verified
- `apply-a-border-of-5-pixels-in-red-color-to-all-extracted-png-images-before-saving.cs`
  - Task: Apply a border of 5 pixels in red color to all extracted PNG images before saving.
  - Workflow: image-extraction
  - Outputs: png
  - Selected engine: verified
- `batch-extract-images-from-a-collection-of-doc-files-and-generate-an-excel-index-with-file.cs`
  - Task: Batch extract images from a collection of DOC files and generate an Excel index with file paths.
  - Workflow: image-extraction
  - Outputs: doc
  - Selected engine: verified
- `convert-extracted-tiff-images-to-grayscale-jpeg-for-use-in-low-bandwidth-environments.cs`
  - Task: Convert extracted TIFF images to grayscale JPEG for use in low‑bandwidth environments.
  - Workflow: image-extraction
  - Outputs: jpeg, tiff
  - Selected engine: verified
- `load-a-docx-template-programmatically-insert-images-from-a-rest-api-and-save-as-pdf.cs`
  - Task: Load a DOCX template, programmatically insert images from a REST API, and save as PDF.
  - Workflow: image-insertion-and-shape
  - Outputs: docx, pdf
  - Selected engine: verified
- `extract-images-from-a-word-document-s-embedded-chart-data-and-export-them-as-svg-files.cs`
  - Task: Extract images from a Word document's embedded chart data and export them as SVG files.
  - Workflow: image-extraction
  - Outputs: docx, svg
  - Selected engine: verified
- `batch-process-docx-files-extracting-images-and-creating-a-summary-csv-containing-image-siz.cs`
  - Task: Batch process DOCX files, extracting images and creating a summary CSV containing image sizes.
  - Workflow: image-extraction
  - Outputs: docx, csv
  - Selected engine: verified
- `resize-extracted-png-images-to-75-of-original-dimensions-for-quick-preview-generation.cs`
  - Task: Resize extracted PNG images to 75% of original dimensions for quick preview generation.
  - Workflow: image-extraction
  - Outputs: png
  - Selected engine: verified
- `apply-a-gaussian-blur-to-all-extracted-jpeg-images-before-embedding-them-into-a-new-docume.cs`
  - Task: Apply a Gaussian blur to all extracted JPEG images before embedding them into a new document.
  - Workflow: image-extraction
  - Outputs: jpeg
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-embedded-3d-model-thumbnails-and-save-as-png.cs`
  - Task: Extract images from a DOCX file's embedded 3D model thumbnails and save as PNG.
  - Workflow: image-extraction
  - Outputs: docx, png
  - Selected engine: verified
- `batch-convert-extracted-gif-animations-to-mp4-video-clips-while-preserving-frame-order.cs`
  - Task: Batch convert extracted GIF animations to MP4 video clips while preserving frame order.
  - Workflow: image-extraction
  - Outputs: gif
  - Selected engine: verified
- `load-a-word-document-extract-images-and-generate-a-latex-file-referencing-each-image.cs`
  - Task: Load a Word document, extract images, and generate a LaTeX file referencing each image.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-custom-xml-data-and-map-them-to-database-records.cs`
  - Task: Extract images from a DOCX file's custom XML data and map them to database records.
  - Workflow: image-extraction
  - Outputs: docx, xml
  - Selected engine: verified
- `resize-extracted-bmp-images-to-640x480-resolution-for-use-in-legacy-applications.cs`
  - Task: Resize extracted BMP images to 640x480 resolution for use in legacy applications.
  - Workflow: image-extraction
  - Outputs: bmp
  - Selected engine: verified
- `apply-a-hue-rotation-of-180-degrees-to-all-extracted-png-images-for-color-testing.cs`
  - Task: Apply a hue rotation of 180 degrees to all extracted PNG images for color testing.
  - Workflow: image-extraction
  - Outputs: png
  - Selected engine: verified
- `batch-extract-images-from-a-set-of-odt-files-and-create-a-searchable-pdf-catalog.cs`
  - Task: Batch extract images from a set of ODT files and create a searchable PDF catalog.
  - Workflow: image-extraction
  - Outputs: pdf
  - Selected engine: verified
- `convert-extracted-jpeg-images-to-high-quality-webp-for-optimized-web-delivery.cs`
  - Task: Convert extracted JPEG images to high‑quality WebP for optimized web delivery.
  - Workflow: image-extraction
  - Outputs: jpeg, webp
  - Selected engine: verified
- `load-a-docx-file-replace-all-images-with-placeholders-and-export-to-html-preserving-layout.cs`
  - Task: Load a DOCX file, replace all images with placeholders, and export to HTML preserving layout.
  - Workflow: general-image-workflow
  - Outputs: docx, html
  - Selected engine: verified
- `extract-images-from-a-word-document-s-embedded-audio-tracks-and-save-them-as-jpeg-thumbnai.cs`
  - Task: Extract images from a Word document's embedded audio tracks and save them as JPEG thumbnails.
  - Workflow: image-extraction
  - Outputs: docx, jpeg
  - Selected engine: verified
- `batch-process-multiple-doc-files-extracting-images-and-generating-a-consolidated-pdf-repor.cs`
  - Task: Batch process multiple DOC files, extracting images and generating a consolidated PDF report.
  - Workflow: image-extraction
  - Outputs: doc, pdf
  - Selected engine: verified
- `resize-extracted-png-images-to-a-maximum-dimension-of-1200-pixels-maintaining-aspect-ratio.cs`
  - Task: Resize extracted PNG images to a maximum dimension of 1200 pixels, maintaining aspect ratio.
  - Workflow: image-extraction
  - Outputs: png
  - Selected engine: verified
- `apply-a-motion-blur-effect-to-all-extracted-jpeg-images-before-re-embedding-them.cs`
  - Task: Apply a motion blur effect to all extracted JPEG images before re‑embedding them.
  - Workflow: image-extraction
  - Outputs: jpeg
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-embedded-spreadsheet-charts-and-save-them-as-png-files.cs`
  - Task: Extract images from a DOCX file's embedded spreadsheet charts and save them as PNG files.
  - Workflow: image-extraction
  - Outputs: docx, png
  - Selected engine: verified
- `batch-convert-extracted-bmp-images-to-lossless-png-and-generate-a-checksum-file-for-verifi.cs`
  - Task: Batch convert extracted BMP images to lossless PNG and generate a checksum file for verification.
  - Workflow: image-extraction
  - Outputs: png, bmp
  - Selected engine: verified
- `load-a-word-document-extract-images-and-create-a-markdown-file-with-image-links.cs`
  - Task: Load a Word document, extract images, and create a Markdown file with image links.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-embedded-video-frames-and-store-them-as-high-resolution.cs`
  - Task: Extract images from a DOCX file's embedded video frames and store them as high‑resolution PNG.
  - Workflow: image-extraction
  - Outputs: docx, png
  - Selected engine: verified
- `resize-extracted-jpeg-images-to-50-width-and-height-for-thumbnail-generation-in-galleries.cs`
  - Task: Resize extracted JPEG images to 50% width and height for thumbnail generation in galleries.
  - Workflow: image-extraction
  - Outputs: jpeg
  - Selected engine: verified
- `apply-a-color-balance-adjustment-to-all-extracted-png-images-before-saving-them-to-the-out.cs`
  - Task: Apply a color balance adjustment to all extracted PNG images before saving them to the output folder.
  - Workflow: image-extraction
  - Outputs: png
  - Selected engine: verified
- `batch-extract-images-from-a-collection-of-docx-files-and-generate-an-html-index-page.cs`
  - Task: Batch extract images from a collection of DOCX files and generate an HTML index page.
  - Workflow: image-extraction
  - Outputs: docx, html
  - Selected engine: verified
- `convert-extracted-tiff-images-to-pdf-with-each-image-on-a-separate-page-and-embed-metadata.cs`
  - Task: Convert extracted TIFF images to PDF with each image on a separate page and embed metadata.
  - Workflow: image-extraction
  - Outputs: pdf, tiff
  - Selected engine: verified
- `load-a-docx-template-programmatically-insert-images-from-a-file-system-and-save-as-docx.cs`
  - Task: Load a DOCX template, programmatically insert images from a file system, and save as DOCX.
  - Workflow: image-insertion-and-shape
  - Outputs: docx, doc
  - Selected engine: verified
- `extract-images-from-a-word-document-s-embedded-qr-codes-and-decode-them-for-data-extractio.cs`
  - Task: Extract images from a Word document's embedded QR codes and decode them for data extraction.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified
- `batch-process-doc-files-extracting-images-and-creating-a-compressed-zip-archive-with-passw.cs`
  - Task: Batch process DOC files, extracting images and creating a compressed ZIP archive with password protection.
  - Workflow: image-extraction
  - Outputs: doc
  - Selected engine: verified
- `resize-extracted-png-images-to-300x300-pixels-and-add-a-watermark-text-overlay.cs`
  - Task: Resize extracted PNG images to 300x300 pixels and add a watermark text overlay.
  - Workflow: image-extraction
  - Outputs: png
  - Selected engine: verified
- `apply-a-saturation-increase-of-20-to-all-extracted-jpeg-images-before-re-saving-them.cs`
  - Task: Apply a saturation increase of 20% to all extracted JPEG images before re‑saving them.
  - Workflow: image-extraction
  - Outputs: jpeg
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-embedded-audio-waveform-visualizations-and-save-as-png.cs`
  - Task: Extract images from a DOCX file's embedded audio waveform visualizations and save as PNG.
  - Workflow: image-extraction
  - Outputs: docx, png
  - Selected engine: verified
- `batch-convert-extracted-gif-images-to-animated-png-while-preserving-original-frame-timing.cs`
  - Task: Batch convert extracted GIF images to animated PNG while preserving original frame timing.
  - Workflow: image-extraction
  - Outputs: png, gif
  - Selected engine: verified
- `load-a-word-document-extract-images-and-generate-a-json-file-containing-base64-encoded-dat.cs`
  - Task: Load a Word document, extract images, and generate a JSON file containing base64 encoded data.
  - Workflow: image-extraction
  - Outputs: docx, json
  - Selected engine: verified
- `extract-images-from-a-docx-file-s-embedded-map-tiles-and-store-them-using-tile-coordinates.cs`
  - Task: Extract images from a DOCX file's embedded map tiles and store them using tile coordinates.
  - Workflow: image-extraction
  - Outputs: docx
  - Selected engine: verified

## Common failure patterns and preferred agent fixes

- **Missing local image bootstrap**
  - Symptom: Examples assume an existing image file or source document already exists.
  - Preferred fix: Create deterministic local sample images and source documents before insertion, extraction, or conversion.

- **Shape insertion at invalid location**
  - Symptom: Runtime failures such as invalid insertion location or self-append mistakes.
  - Preferred fix: Append image shapes only once and only into valid containers such as Paragraph.

- **Unsafe image extraction**
  - Symptom: Code accesses ImageData on nodes that are not image-bearing shapes.
  - Preferred fix: Enumerate Shape nodes and require shape.HasImage before accessing or saving ImageData.

- **Stream reuse bugs**
  - Symptom: Image or document stream operations fail because MemoryStream.Position was not reset.
  - Preferred fix: Always reset stream.Position = 0 before reuse.

- **Context-specific image container mistakes**
  - Symptom: Comment, content-control, footnote, or property-based image tasks fail due to wrong document structure assumptions.
  - Preferred fix: Traverse the actual container nodes first, then shapes inside them, and insert shapes only through valid child containers.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`
- Additional package: `Aspose.Drawing.Common`
- Additional package: `Newtonsoft.Json`

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
dotnet add package Newtonsoft.Json
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\images\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer deterministic local image assets over external dependencies.
