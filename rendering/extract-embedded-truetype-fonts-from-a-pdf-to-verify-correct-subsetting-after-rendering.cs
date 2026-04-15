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

        // Create a simple document that uses a TrueType font (Arial).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This text is rendered with Arial to test font subsetting.");

        // Configure PDF save options to enable subsetting (default behavior).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure subsetting is active (do not embed full fonts).
            EmbedFullFonts = false,
            // Embed all fonts; subsetting will still be applied.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF.
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        doc.Save(pdfPath, pdfOptions);

        // Verify that the generated PDF contains markers of embedded TrueType fonts.
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfContent = Encoding.ASCII.GetString(pdfBytes);

        bool containsSubsetFontName = Regex.IsMatch(pdfContent, @"[A-Z]{6}\+");
        bool containsFontFileMarkers = pdfContent.Contains("/FontFile") ||
                                       pdfContent.Contains("/FontFile2") ||
                                       pdfContent.Contains("/FontFile3") ||
                                       pdfContent.Contains("/Subtype /TrueType");

        if (!containsSubsetFontName && !containsFontFileMarkers)
        {
            throw new Exception("Verification failed: No embedded TrueType font markers were found in the PDF.");
        }

        // Indicate successful verification.
        Console.WriteLine("Embedded TrueType font markers verified successfully.");
    }
}
