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
            // Folder that contains the source ODT files.
            string odtFolder = @"C:\InputOdt";

            // Folder where extracted images will be saved.
            string imagesFolder = @"C:\ExtractedImages";

            // Path of the final searchable PDF catalog.
            string catalogPdfPath = @"C:\Catalog.pdf";

            // Ensure the images folder exists.
            if (!Directory.Exists(imagesFolder))
                Directory.CreateDirectory(imagesFolder);

            // Create a blank document that will become the PDF catalog.
            Document catalogDoc = new Document(); // uses Document() constructor rule
            DocumentBuilder catalogBuilder = new DocumentBuilder(catalogDoc);

            // Process each ODT file in the source folder.
            foreach (string odtFilePath in Directory.GetFiles(odtFolder, "*.odt"))
            {
                // Load the ODT document (Document(string) constructor rule).
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

                    // Save the image to the file system (ImageData.Save method).
                    shape.ImageData.Save(imageFullPath);
                    imageCounter++;

                    // Add a heading for the image in the catalog.
                    catalogBuilder.InsertParagraph();
                    catalogBuilder.Writeln($"Image extracted from: {Path.GetFileName(odtFilePath)}");

                    // Insert the extracted image into the catalog document.
                    // InsertImage(string) inserts the image inline at 100% scale.
                    catalogBuilder.InsertImage(imageFullPath);
                }
            }

            // Configure PDF save options to produce a searchable PDF (PDF/A‑2u compliance).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA2u // ensures text is searchable and document is archivable
            };

            // Save the catalog document as PDF using the Save(string, SaveOptions) overload.
            catalogDoc.Save(catalogPdfPath, pdfOptions);
        }
    }
}
