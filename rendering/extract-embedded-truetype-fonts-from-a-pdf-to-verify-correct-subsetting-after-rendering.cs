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
        // Define folders for output.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the generated PDF.
        string pdfPath = Path.Combine(outputDir, "Sample.pdf");

        // Create a simple document that uses a TrueType font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Common TrueType font.
        builder.Writeln("This text is rendered with Arial to test font subsetting.");

        // Configure PDF save options to embed all fonts (subsetting will be applied by default).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll,
            EmbedFullFonts = false // Ensure subsetting, not full embedding.
        };

        // Save the document as PDF.
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Load the PDF content as a string for simple text inspection.
        // PDF is binary; using ISO-8859-1 preserves byte values in the string.
        string pdfContent = File.ReadAllText(pdfPath, Encoding.GetEncoding("ISO-8859-1"));

        // Look for typical markers of embedded TrueType fonts.
        bool hasFontFileMarker = pdfContent.Contains("/FontFile") ||
                                 pdfContent.Contains("/FontFile2") ||
                                 pdfContent.Contains("/FontFile3");

        // Look for subset font name pattern: six uppercase letters followed by '+' (e.g., ABCDEF+Arial).
        bool hasSubsetFontName = Regex.IsMatch(pdfContent, @"[A-Z]{6}\+");

        // Validate that at least one of the expected markers is present.
        if (!hasFontFileMarker || !hasSubsetFontName)
            throw new InvalidOperationException("Embedded subset TrueType font markers were not found in the PDF.");

        // If we reach this point, verification succeeded.
        Console.WriteLine("PDF generated successfully and contains embedded subset TrueType font markers.");
    }
}
