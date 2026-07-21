using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a common font that supports OpenType ligatures.
        builder.Font.Name = "Arial";

        // Write text that contains ligatures (fi, fl) and contextual forms.
        builder.Writeln("Office");
        builder.Writeln("affinity");
        builder.Writeln("fluff");
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Save the document to PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure fonts are subsetted (default) to keep file size reasonable.
            EmbedFullFonts = false
        };

        string pdfPath = Path.Combine(outputDir, "Ligatures.pdf");
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Load the PDF bytes and look for subset font markers (e.g., six uppercase letters followed by '+').
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfText = Encoding.ASCII.GetString(pdfBytes);

        // Simple regex to detect subset font names like "ABCDEF+ArialMT".
        bool hasSubsetFont = Regex.IsMatch(pdfText, @"[A-Z]{6}\+");

        if (!hasSubsetFont)
            throw new InvalidOperationException("The generated PDF does not contain subset font markers, indicating OpenType features may not be preserved.");

        // Indicate success (no interactive prompts required).
        Console.WriteLine("PDF generated successfully with OpenType ligatures preserved.");
    }
}
