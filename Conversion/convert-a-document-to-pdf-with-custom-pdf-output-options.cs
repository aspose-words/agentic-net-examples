using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class PdfConverter
{
    /// <summary>
    /// Converts a Word document to PDF applying custom PDF save options.
    /// </summary>
    /// <param name="inputPath">Full path to the source .doc/.docx file.</param>
    /// <param name="outputPath">Full path where the resulting PDF will be saved.</param>
    public void ConvertToPdf(string inputPath, string outputPath)
    {
        // Load the source document.
        Document doc = new Document(inputPath);

        // Create a PdfSaveOptions instance via the factory method.
        PdfSaveOptions pdfOptions = (PdfSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

        // ----- Custom PDF output options -----
        // Example: set PDF/A-2u compliance.
        pdfOptions.Compliance = PdfCompliance.PdfA2u;

        // Export custom document properties as standard entries in the PDF info dictionary.
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Enable high‑quality rendering (slower but better visual fidelity).
        pdfOptions.UseHighQualityRendering = true;

        // Show the document outline (bookmarks) when the PDF is opened.
        pdfOptions.PageMode = PdfPageMode.UseOutlines;

        // Save the document as PDF using the configured options.
        doc.Save(outputPath, pdfOptions);
    }
}

public static class Program
{
    /// <summary>
    /// Entry point required for a console application.
    /// </summary>
    /// <param name="args">[0] – input .doc/.docx path, [1] – output .pdf path.</param>
    public static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: PdfConverter <input.docx> <output.pdf>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            var converter = new PdfConverter();
            converter.ConvertToPdf(inputPath, outputPath);
            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}
