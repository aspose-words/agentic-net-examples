using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new document and add some text that would normally use OpenType ligatures.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 48;
        builder.Writeln("Office");          // Contains "ff" ligature.
        builder.Writeln("fi fl ffi ffl");   // Contains "fi", "fl", "ffi", "ffl" ligatures.

        // Disable OpenType font formatting features for the whole document.
        doc.CompatibilityOptions.DisableOpenTypeFontFormattingFeatures = true;

        // Save the document to PDF.
        string pdfPath = Path.Combine(outputDir, "DisabledOpenType.pdf");
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        doc.Save(pdfPath, saveOptions);

        // Verify that the PDF file was created and is not empty.
        if (!File.Exists(pdfPath))
            throw new Exception("PDF file was not created.");

        if (new FileInfo(pdfPath).Length == 0)
            throw new Exception("PDF file is empty.");
    }
}
