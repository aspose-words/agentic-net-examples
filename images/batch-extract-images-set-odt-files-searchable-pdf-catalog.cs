using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

namespace OdtImageCatalog
{
    class Program
    {
        static void Main()
        {
            // Base directory (the folder where the executable runs).
            string baseDir = AppContext.BaseDirectory;

            // Folder that contains the source ODT files.
            string odtFolder = Path.Combine(baseDir, "InputOdt");

            // Folder where extracted images will be saved.
            string imagesFolder = Path.Combine(baseDir, "ExtractedImages");

            // Path of the final searchable PDF catalog.
            string catalogPdfPath = Path.Combine(baseDir, "Catalog.pdf");

            // Ensure the required folders exist.
            Directory.CreateDirectory(odtFolder);
            Directory.CreateDirectory(imagesFolder);

            // Create a blank document that will become the PDF catalog.
            Document catalogDoc = new Document();
            DocumentBuilder catalogBuilder = new DocumentBuilder(catalogDoc);

            // Get all ODT files in the source folder.
            string[] odtFiles = Directory.GetFiles(odtFolder, "*.odt");
            if (odtFiles.Length == 0)
            {
                Console.WriteLine($"No ODT files found in '{odtFolder}'. Place files there and rerun the program.");
                return;
            }

            // Process each ODT file.
            foreach (string odtFilePath in odtFiles)
            {
                // Load the ODT document.
                Document sourceDoc = new Document(odtFilePath);

                // Collect all Shape nodes that may contain images.
                NodeCollection shapeNodes = sourceDoc.GetChildNodes(NodeType.Shape, true);

                int imageCounter = 0;

                foreach (Shape shape in shapeNodes.OfType<Shape>())
                {
                    if (!shape.HasImage)
                        continue;

                    // Build a unique file name for the extracted image.
                    string imageExtension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(odtFilePath)}_{imageCounter}{imageExtension}";
                    string imageFullPath = Path.Combine(imagesFolder, imageFileName);

                    // Save the image to the file system.
                    shape.ImageData.Save(imageFullPath);
                    imageCounter++;

                    // Add a heading for the image in the catalog.
                    catalogBuilder.InsertParagraph();
                    catalogBuilder.Writeln($"Image extracted from: {Path.GetFileName(odtFilePath)}");

                    // Insert the extracted image into the catalog document.
                    catalogBuilder.InsertImage(imageFullPath);
                }
            }

            // Configure PDF save options to produce a searchable PDF (PDF/A‑2u compliance).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA2u
            };

            // Save the catalog document as PDF.
            catalogDoc.Save(catalogPdfPath, pdfOptions);

            Console.WriteLine($"Catalog PDF created at: {catalogPdfPath}");
        }
    }
}
