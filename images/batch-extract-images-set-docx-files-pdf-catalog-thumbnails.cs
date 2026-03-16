using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ImageCatalogGenerator
{
    static void Main()
    {
        // Folder containing the source DOCX files.
        string sourceFolder = @"C:\Docs\Input";
        // Path for the resulting PDF catalog.
        string outputPdf = @"C:\Docs\Output\Catalog.pdf";

        // Create a new blank document that will become the catalog.
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);

        // Iterate over every DOCX file in the source folder.
        foreach (string docxPath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            // Load the source document.
            Document sourceDoc = new Document(docxPath);

            // Render the first page of the source document to a thumbnail image.
            // Use ImageSaveOptions to control size and resolution.
            ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render only the first page (zero‑based index).
                PageSet = new PageSet(0),
                // Small resolution keeps the thumbnail lightweight.
                Resolution = 72,
                // Optional: limit pixel dimensions.
                ImageSize = new System.Drawing.Size(150, 200)
            };

            // Save the rendered page to a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                sourceDoc.Save(imgStream, imgOptions);
                imgStream.Position = 0; // Reset for reading.

                // Insert a heading with the source file name.
                builder.Writeln(Path.GetFileName(docxPath));

                // Insert the thumbnail image inline.
                builder.InsertImage(imgStream);

                // Add a blank line after each entry for readability.
                builder.Writeln();
            }
        }

        // Save the assembled catalog as a PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Optional: set PDF compliance or other options here.
            // Example: pdfOptions.Compliance = PdfCompliance.PdfA1b;
        };

        catalog.Save(outputPdf, pdfOptions);
    }
}
