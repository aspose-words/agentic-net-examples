using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Expect exactly two arguments: input file and output file.
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: PdfConverter <inputPath> <outputPath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            ConvertToPdf(inputPath, outputPath);
            Console.WriteLine($"Document converted successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error converting document: {ex.Message}");
        }
    }

    /// <summary>
    /// Converts a document from any supported format to PDF.
    /// </summary>
    /// <param name="inputPath">Full path to the source document.</param>
    /// <param name="outputPath">Full path where the PDF will be saved.</param>
    public static void ConvertToPdf(string inputPath, string outputPath)
    {
        // Load the source document using Aspose.Words' built‑in loading mechanism.
        Document document = new Document(inputPath);

        // Create PDF save options – can be customized if needed.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF using the options defined above.
        document.Save(outputPath, pdfOptions);
    }
}
