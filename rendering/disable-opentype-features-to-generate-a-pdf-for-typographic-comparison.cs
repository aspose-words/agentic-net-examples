using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new document and add some text that would normally use OpenType features (e.g., ligatures).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 48;
        builder.Writeln("Office"); // Contains the "ff" ligature in many fonts.

        // Disable OpenType font formatting features for the whole document.
        doc.CompatibilityOptions.DisableOpenTypeFontFormattingFeatures = true;

        // Save the document as PDF.
        string pdfPath = Path.Combine(outputDir, "DisabledOpenType.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The PDF file was not created.", pdfPath);
    }
}
