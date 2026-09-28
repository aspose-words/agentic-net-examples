using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class RenderingExample
{
    public static void Main()
    {
        // Define file names.
        const string sourcePath = "Sample.docx";
        const string pdfPath = "RenderedOutput.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a simple DOCX document locally.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words rendering example!");
        // Save the source DOCX.
        doc.Save(sourcePath);

        // -----------------------------------------------------------------
        // Step 2: Load the DOCX document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // Step 3: Configure rendering options for PDF conversion.
        // -----------------------------------------------------------------
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Example option: set PDF compliance level.
            Compliance = PdfCompliance.PdfA1b,
            // Example option: compress images using JPEG.
            ImageCompression = PdfImageCompression.Jpeg,
            // Example option: set JPEG quality (0-100).
            JpegQuality = 90
        };

        // -----------------------------------------------------------------
        // Step 4: Render the document to PDF using the configured options.
        // -----------------------------------------------------------------
        loadedDoc.Save(pdfPath, saveOptions);

        // -----------------------------------------------------------------
        // Step 5: Validate that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException($"Failed to create the PDF file at '{pdfPath}'.");
        }

        // Optional: Output a simple confirmation.
        Console.WriteLine($"Document rendered successfully to '{pdfPath}'.");
    }
}
