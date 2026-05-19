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

        // Create a simple document that uses a TrueType font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Arial is a TrueType font.
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");

        // Save the document to PDF with default subsetting (EmbedFullFonts = false).
        string pdfPath = Path.Combine(outputDir, "Sample.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions(); // Subsetting is enabled by default.
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF contains embedded TrueType font markers.
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfText = Encoding.ASCII.GetString(pdfBytes);

        bool containsFontFileMarker = pdfText.Contains("/FontFile") ||
                                      pdfText.Contains("/FontFile2") ||
                                      pdfText.Contains("/FontFile3");

        // Subset fonts are usually named with six uppercase letters followed by '+'.
        bool containsSubsetName = Regex.IsMatch(pdfText, @"[A-Z]{6}\+");

        if (!containsFontFileMarker && !containsSubsetName)
        {
            throw new Exception("The generated PDF does not contain embedded TrueType font markers.");
        }

        Console.WriteLine("Embedded TrueType font markers verified successfully.");
    }
}
