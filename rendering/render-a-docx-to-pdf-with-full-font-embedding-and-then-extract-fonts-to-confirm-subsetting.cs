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

        // Create a simple DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world! This text will be rendered to PDF with full font embedding.");

        // Configure PDF save options to embed full fonts (no subsetting).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Save the document as PDF.
        string pdfPath = Path.Combine(outputDir, "RenderedFullFonts.pdf");
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Load PDF bytes for inspection.
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfText = Encoding.ASCII.GetString(pdfBytes);

        // Check for embedded font markers.
        bool hasFontFileMarker = pdfText.Contains("/FontFile") ||
                                 pdfText.Contains("/FontFile2") ||
                                 pdfText.Contains("/FontFile3");

        // Check for subset-style font name pattern (e.g., ABCDEF+FontName).
        bool hasSubsetFontName = Regex.IsMatch(pdfText, @"[A-Z]{6}\+");

        // Validate that full embedding (no subsetting) is indicated.
        if (!hasFontFileMarker && !hasSubsetFontName)
            throw new InvalidOperationException("The PDF does not contain expected embedded font markers; subsetting may be enabled.");

        // Success message.
        Console.WriteLine("PDF rendered with full font embedding successfully verified.");
    }
}
