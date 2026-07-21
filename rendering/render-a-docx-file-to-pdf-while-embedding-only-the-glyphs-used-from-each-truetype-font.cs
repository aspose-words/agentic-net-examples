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
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document that uses a TrueType font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Standard TrueType font.
        builder.Writeln("Hello world! This PDF will contain only the glyphs used from the font.");

        // Configure PDF save options to use subsetting (default behavior).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = false // Ensure only used glyphs are embedded.
        };

        // Save the document as PDF.
        string pdfPath = Path.Combine(artifactsDir, "SampleSubset.pdf");
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Load the PDF bytes and look for subset font markers (e.g., "ABCDEF+Arial").
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfText = Encoding.ASCII.GetString(pdfBytes);

        // Regex pattern for a subset font name: six uppercase letters followed by '+' and the font name.
        string pattern = @"[A-Z]{6}\+Arial";
        bool hasSubsetMarker = Regex.IsMatch(pdfText, pattern);

        if (!hasSubsetMarker)
            throw new InvalidOperationException("The generated PDF does not contain subset font markers, indicating that subsetting may have failed.");

        // Success – the PDF was generated with subsetted TrueType fonts.
        Console.WriteLine("PDF generated successfully with subsetted TrueType fonts at:");
        Console.WriteLine(pdfPath);
    }
}
