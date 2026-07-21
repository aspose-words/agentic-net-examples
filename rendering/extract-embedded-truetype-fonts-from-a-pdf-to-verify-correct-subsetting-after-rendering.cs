using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define output paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "Sample.pdf");

        // Create a simple document that uses a TrueType font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Common TrueType font.
        builder.Writeln("This text will be rendered with Arial and should be subsetted in the PDF.");

        // Configure PDF save options to embed fonts (subsetted).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed all fonts, but keep subsetting (default false for EmbedFullFonts).
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll,
            EmbedFullFonts = false
        };

        // Save the document as PDF.
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Load the PDF content as text for inspection.
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfText = Encoding.ASCII.GetString(pdfBytes);

        // Look for typical font embedding markers.
        bool hasFontFileMarker = pdfText.Contains("/FontFile") ||
                                 pdfText.Contains("/FontFile2") ||
                                 pdfText.Contains("/FontFile3");

        // Look for subset font name pattern (e.g., ABCDEF+Arial).
        bool hasSubsetFontName = Regex.IsMatch(pdfText, @"[A-Z]{6}\+");

        // Validate that at least one of the expected markers is present.
        if (!hasFontFileMarker && !hasSubsetFontName)
            throw new InvalidOperationException("The generated PDF does not contain embedded TrueType font markers.");

        // Output verification result.
        Console.WriteLine("PDF generated successfully with embedded (subsetted) TrueType fonts.");
        Console.WriteLine($"File location: {pdfPath}");
    }
}
