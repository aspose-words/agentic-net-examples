using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "LigaturesPreserved.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that contains common ligatures (e.g., Calibri).
        builder.Font.Name = "Calibri";
        builder.Font.Size = 24;

        // Write a paragraph that includes ligature characters such as "ff", "fi", "fl".
        builder.Writeln("Office offers efficient workflow. The office staff often use the word \"office\" which contains the \"ff\" ligature.");
        builder.Writeln("A fine example of ligatures is the \"fi\" and \"fl\" combinations in this sentence.");

        // Save the document to PDF. PdfSaveOptions is used to control PDF rendering.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Ensure that fonts are embedded (subsetting) so that ligatures are preserved in the PDF.
            EmbedFullFonts = false,
            // Use high‑quality rendering to improve typographic fidelity.
            UseHighQualityRendering = true,
            UseAntiAliasing = true
        };

        doc.Save(pdfPath, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The PDF file was not created.", pdfPath);

        // Optional: output the path for confirmation (no interactive prompt).
        Console.WriteLine($"PDF saved to: {pdfPath}");
    }
}
