using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names for the sample DOCX and the rendered PDF.
        string docPath = "Sample.docx";
        string pdfPath = "RenderedOutput.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a simple DOCX document locally.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words rendering example!");
        // Save the sample document to disk.
        doc.Save(docPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Step 2: Load the DOCX document from the file system.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // Step 3: Configure rendering options before conversion.
        // Example: Save as PDF/A-1b with JPEG image compression.
        // -----------------------------------------------------------------
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b,
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 90
        };

        // -----------------------------------------------------------------
        // Step 4: Render the document to PDF using the configured options.
        // -----------------------------------------------------------------
        loadedDoc.Save(pdfPath, saveOptions);

        // -----------------------------------------------------------------
        // Step 5: Validate that the PDF file was created successfully.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException($"Failed to create the PDF file at '{pdfPath}'.");
        }

        // Optional: Output file size for confirmation (no user interaction required).
        long fileSize = new FileInfo(pdfPath).Length;
        Console.WriteLine($"PDF rendered successfully. File size: {fileSize} bytes.");
    }
}
