using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample DOCX and the rendered PDF.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string docxPath = Path.Combine(artifactsDir, "SampleDocument.docx");
        string pdfPath = Path.Combine(artifactsDir, "RenderedDocument.pdf");

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document if it does not already exist.
        // -----------------------------------------------------------------
        if (!File.Exists(docxPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("Hello Aspose.Words!");
            builder.Writeln("This document will be rendered to PDF with custom options.");
            sampleDoc.Save(docxPath);
        }

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);

        // -----------------------------------------------------------------
        // 3. Configure rendering options before conversion.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: render the PDF in grayscale.
            ColorMode = ColorMode.Grayscale,

            // Example: use high‑quality rendering (slower but better visual fidelity).
            UseHighQualityRendering = true,

            // Example: embed fonts as subsets to keep file size small.
            EmbedFullFonts = false,

            // Example: render DrawingML shapes as they are (no fallback).
            DmlRenderingMode = DmlRenderingMode.DrawingML
        };

        // -----------------------------------------------------------------
        // 4. Convert the document to PDF using the configured options.
        // -----------------------------------------------------------------
        loadedDoc.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 5. Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional: indicate success.
        Console.WriteLine("Document rendered successfully to: " + pdfPath);
    }
}
