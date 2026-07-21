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
        string pdfPath = Path.Combine(outputDir, "TypographicComparison.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that contains OpenType features (e.g., ligatures) and write sample text.
        builder.Font.Name = "Arial";
        builder.Font.Size = 48;
        builder.Writeln("Office"); // Contains the "fi" ligature.

        // Disable OpenType font formatting features for the whole document.
        doc.CompatibilityOptions.DisableOpenTypeFontFormattingFeatures = true;

        // Save the document as PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF file.");

        Console.WriteLine($"PDF generated successfully at: {pdfPath}");
    }
}
