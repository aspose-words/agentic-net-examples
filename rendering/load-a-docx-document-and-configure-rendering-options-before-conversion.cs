using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Path for the sample DOCX file.
        string sourceDocPath = Path.Combine(artifactsDir, "Sample.docx");

        // Create a simple DOCX document if it does not already exist.
        if (!File.Exists(sourceDocPath))
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello Aspose.Words!");
            doc.Save(sourceDocPath);
        }

        // Load the DOCX document.
        Document loadedDoc = new Document(sourceDocPath);

        // Configure rendering options for PDF conversion.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed the full font data (no subsetting).
            EmbedFullFonts = true,
            // Use high‑quality rendering algorithms.
            UseHighQualityRendering = true,
            // Render colors in grayscale.
            ColorMode = ColorMode.Grayscale
        };

        // Path for the output PDF file.
        string pdfPath = Path.Combine(artifactsDir, "Sample.pdf");

        // Save the document as PDF using the configured options.
        loadedDoc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new Exception("Failed to create the PDF file.");

        Console.WriteLine("PDF successfully created at: " + pdfPath);
    }
}
