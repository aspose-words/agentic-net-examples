using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple DOCX document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Font.Size = 24;
        builder.Writeln("This is a sample document created for rendering demonstration.");

        // Configure PDF rendering options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use high‑quality rendering algorithms.
            UseHighQualityRendering = true,
            // Do not embed full fonts (use subsetting to keep file size small).
            EmbedFullFonts = false,
            // Render colors normally.
            ColorMode = ColorMode.Normal
        };

        // Save the document as PDF using the configured options.
        string pdfPath = Path.Combine(outputDir, "RenderedDocument.pdf");
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output file.");

        // Optionally, inform that the process completed successfully.
        Console.WriteLine($"PDF successfully saved to: {pdfPath}");
    }
}
