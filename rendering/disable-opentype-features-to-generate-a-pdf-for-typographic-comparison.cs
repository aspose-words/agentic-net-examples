using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "DisabledOpenType.pdf");

        // Create a new document and add some text that may use OpenType features (e.g., ligatures).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("Office"); // Contains the "ff" ligature.

        // Disable OpenType font formatting features for the whole document.
        doc.CompatibilityOptions.DisableOpenTypeFontFormattingFeatures = true;

        // Save the document to PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF file.");

        // Indicate successful completion (no interactive prompts).
        Console.WriteLine("PDF generated at: " + pdfPath);
    }
}
