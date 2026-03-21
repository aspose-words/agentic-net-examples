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
            // Ensure the input folder exists.
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder \"{inputFolder}\" does not exist. No files will be processed.");
                return;
            }

            // Ensure the images output directory exists.
            if (!Directory.Exists(imagesFolder))
                Directory.CreateDirectory(imagesFolder);

            // Collect information about each extracted image for the catalog.
            var catalogEntries = new List<(string SourceDocument, string ImagePath)>();

            // Load each DOCX file, extract its images, and store the mapping.
            foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
            {
                // Load the source document.
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

                    // Save the image data to disk.
                    shape.ImageData.Save(imageFullPath);

                    // Record the entry for later inclusion in the PDF catalog.
                    catalogEntries.Add((SourceDocument: sourceFileName, ImagePath: imageFullPath));

                    imageIndex++;
                }
            }

            // Create a new blank document that will become the PDF catalog.
            Document catalogDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(catalogDoc);

            // Title page.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
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
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Font.Size = 12;
                builder.Font.Bold = true;
                builder.Writeln($"Source Document: {entry.SourceDocument}");

                builder.InsertImage(entry.ImagePath);
                builder.Writeln();
                builder.InsertBreak(BreakType.PageBreak);
            }

            // Save the catalog as PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.Pdf17
            };
            catalogDoc.Save(catalogPdfPath, pdfOptions);
        }
    }

    class Program
    {
        static void Main()
        {
            // Use paths relative to the executable's working directory.
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string inputFolder = Path.Combine(baseDir, "Input");
            string imagesFolder = Path.Combine(baseDir, "ExtractedImages");
            string catalogPdf = Path.Combine(baseDir, "ImageCatalog.pdf");

            // Ensure the required directories exist.
            Directory.CreateDirectory(inputFolder);
            Directory.CreateDirectory(imagesFolder);

            Console.WriteLine($"Input folder: {inputFolder}");
            Console.WriteLine($"Images folder: {imagesFolder}");
            Console.WriteLine($"Catalog PDF will be saved to: {catalogPdf}");

            ImageCatalogGenerator.Generate(inputFolder, imagesFolder, catalogPdf);

            Console.WriteLine("Processing complete.");
        }
    }
}
