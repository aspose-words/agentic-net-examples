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

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document locally.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello Aspose.Words rendering!");
        builder.Writeln("This document will be converted with custom rendering options.");
        sampleDoc.Save(docPath); // Save the source DOCX.

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // 3. Configure rendering options before conversion.
        //    Example: render to PDF with grayscale images and high‑quality rendering.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Render all images in grayscale.
            ColorMode = ColorMode.Grayscale,
            // Use high‑quality (slower) rendering algorithms.
            UseHighQualityRendering = true,
            // Do not embed full fonts; use subsetting to keep file size small.
            EmbedFullFonts = false
        };

        // -----------------------------------------------------------------
        // 4. Save the document to PDF using the configured options.
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(artifactsDir, "Rendered.pdf");
        loadedDoc.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 5. Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        Console.WriteLine("Document successfully rendered to PDF:");
        Console.WriteLine(pdfPath);
    }
}
