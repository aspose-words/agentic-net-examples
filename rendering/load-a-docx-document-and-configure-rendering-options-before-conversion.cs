using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output files.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document locally.
        // -----------------------------------------------------------------
        string sourceDocPath = Path.Combine(artifactsDir, "SampleDocument.docx");
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello Aspose.Words!");
        builder.Writeln("This document will be converted to PDF with custom rendering options.");
        sampleDoc.Save(sourceDocPath);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourceDocPath);

        // -----------------------------------------------------------------
        // 3. Configure rendering options before conversion.
        //    Here we use PdfSaveOptions to control PDF rendering.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Render with higher quality (slower) algorithms.
            UseHighQualityRendering = true,
            // Reduce memory consumption for large documents.
            MemoryOptimization = true,
            // Example: embed only the glyphs used in the document (subsetting).
            // This property is true by default; setting it explicitly for clarity.
            EmbedFullFonts = false
        };

        // -----------------------------------------------------------------
        // 4. Convert and save the document to PDF using the configured options.
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(artifactsDir, "ConvertedDocument.pdf");
        loadedDoc.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF conversion failed: output file not found.");

        // Optional: write a simple confirmation to the console.
        Console.WriteLine("Document successfully converted to PDF:");
        Console.WriteLine(pdfPath);
    }
}
