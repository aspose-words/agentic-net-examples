using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace AsposeWordsBatchProcessing
{
    /// <summary>
    /// Provides functionality to extract images from a batch of DOCX files
    /// and generate a PDF catalog that lists those images.
    /// </summary>
    public static class ImageCatalogGenerator
    {
        /// <summary>
        /// Processes all DOCX files in <paramref name="inputFolder"/>, extracts every image
        /// to <paramref name="imagesFolder"/>, and creates a PDF catalog at <paramref name="catalogPdfPath"/>.
        /// </summary>
        /// <param name="inputFolder">Folder containing source DOCX files.</param>
        /// <param name="imagesFolder">Folder where extracted images will be saved.</param>
        /// <param name="catalogPdfPath">Full file name of the resulting PDF catalog.</param>
        public static void Generate(string inputFolder, string imagesFolder, string catalogPdfPath)
        {
            // Ensure the images output directory exists.
            if (!Directory.Exists(imagesFolder))
                Directory.CreateDirectory(imagesFolder);

            // Collect information about each extracted image for the catalog.
            var catalogEntries = new List<(string SourceDocument, string ImagePath)>();

            // Load each DOCX file, extract its images, and store the mapping.
            foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
            {
                // Load the source document (lifecycle: load).
                Document sourceDoc = new Document(docxPath);

                // Retrieve all Shape nodes (including images) from the document.
                NodeCollection shapeNodes = sourceDoc.GetChildNodes(NodeType.Shape, true);

                int imageIndex = 0;
                foreach (Shape shape in shapeNodes.OfType<Shape>())
                {
                    // Only process shapes that actually contain an image.
                    if (!shape.HasImage)
                        continue;

                    // Determine a suitable file extension for the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                    // Build a unique file name: <sourceDocName>_img<index>.<ext>
                    string sourceFileName = Path.GetFileNameWithoutExtension(docxPath);
                    string imageFileName = $"{sourceFileName}_img{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(imagesFolder, imageFileName);

                    // Save the image data to disk (lifecycle: save).
                    shape.ImageData.Save(imageFullPath);

                    // Record the entry for later inclusion in the PDF catalog.
                    catalogEntries.Add((SourceDocument: sourceFileName, ImagePath: imageFullPath));

                    imageIndex++;
                }
            }

            // Create a new blank document that will become the PDF catalog (lifecycle: create).
            Document catalogDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(catalogDoc);

            // Optional: add a title page.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center; // Fixed enum
            builder.Font.Size = 24;
            builder.Font.Bold = true;
            builder.Writeln("Image Catalog");
            builder.Font.Size = 12;
            builder.Font.Bold = false;
            builder.Writeln($"Generated on {DateTime.Now}");
            builder.InsertBreak(BreakType.PageBreak);

            // Insert each extracted image with a caption indicating its source document.
            foreach (var entry in catalogEntries)
            {
                // Caption.
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left; // Fixed enum
                builder.Font.Size = 12;
                builder.Font.Bold = true;
                builder.Writeln($"Source Document: {entry.SourceDocument}");

                // Insert the image.
                builder.InsertImage(entry.ImagePath);

                // Add a small spacing after each image.
                builder.Writeln();
                builder.InsertBreak(BreakType.PageBreak);
            }

            // Save the catalog as PDF using the built‑in Save method (lifecycle: save).
            // PdfSaveOptions can be used if additional PDF settings are required.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: set compliance level; adjust as needed.
                Compliance = PdfCompliance.Pdf17
            };
            catalogDoc.Save(catalogPdfPath, pdfOptions);
        }
    }

    // Example usage.
    class Program
    {
        static void Main()
        {
            string inputFolder = @"C:\Docs\Input";          // Folder with DOCX files.
            string imagesFolder = @"C:\Docs\ExtractedImages"; // Where images will be saved.
            string catalogPdf = @"C:\Docs\ImageCatalog.pdf"; // Output PDF catalog.

            ImageCatalogGenerator.Generate(inputFolder, imagesFolder, catalogPdf);

            Console.WriteLine("Processing complete.");
        }
    }
}
