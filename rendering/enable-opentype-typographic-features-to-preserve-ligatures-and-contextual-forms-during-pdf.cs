using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define the folder where the output PDF will be saved.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Choose a font that supports OpenType ligatures (e.g., Calibri).
        builder.Font.Name = "Calibri";
        builder.Font.Size = 24;

        // Write text that contains ligatures and contextual forms.
        builder.Writeln("Office");                     // Contains the "ff" ligature.
        builder.Writeln("efficient");                  // Contains the "fi" ligature.
        builder.Writeln("ﬂ (fl ligature) and ﬁ (fi ligature) demonstration.");

        // Save the document as PDF. The default rendering preserves OpenType features when possible.
        string pdfPath = Path.Combine(artifactsDir, "Ligatures.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        Console.WriteLine($"PDF saved to: {pdfPath}");
    }
}
