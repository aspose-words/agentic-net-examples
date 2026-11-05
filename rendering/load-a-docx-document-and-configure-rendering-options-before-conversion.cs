using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class RenderingExample
{
    public static void Main()
    {
        // Define file paths in the temporary directory.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeWordsRendering");
        Directory.CreateDirectory(tempDir);
        string docxPath = Path.Combine(tempDir, "Sample.docx");
        string pdfPath = Path.Combine(tempDir, "Sample.pdf");

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document locally.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World! This is a sample document for rendering.");
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);

        // -----------------------------------------------------------------
        // 3. Configure PDF rendering options.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: produce PDF/A-1b compliant file.
            Compliance = PdfCompliance.PdfA1b,
            // Example: compress images using JPEG with quality 80.
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 80
        };

        // -----------------------------------------------------------------
        // 4. Render (save) the document to PDF using the configured options.
        // -----------------------------------------------------------------
        loadedDoc.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("PDF rendering failed: output file not found.");
        }

        // Optional: output the location of the generated PDF.
        Console.WriteLine($"PDF successfully created at: {pdfPath}");
    }
}
