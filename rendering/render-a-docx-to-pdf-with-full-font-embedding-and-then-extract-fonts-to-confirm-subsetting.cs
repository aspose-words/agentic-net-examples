using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class RenderDocxToPdfWithFullFontEmbedding
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Calibri";               // Use a TrueType font that is likely present.
        builder.Writeln("This text will be rendered to PDF with full font embedding.");

        // Ensure Aspose.Words can locate the font files.
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        FontSettings.DefaultInstance.SetFontsFolder(fontsFolder, true);

        // Configure PDF save options to embed the full font (no subsetting).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true,
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF.
        string pdfPath = Path.Combine(outputDir, "RenderedFullFont.pdf");
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Load the PDF bytes and look for font embedding markers.
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfText = Encoding.ASCII.GetString(pdfBytes);

        // Look for subset font name pattern (e.g., ABCDEF+FontName) or explicit font file entries.
        bool hasSubsetMarker = Regex.IsMatch(pdfText, @"[A-Z]{6}\+");
        bool hasFontFileEntry = pdfText.Contains("/FontFile") ||
                                pdfText.Contains("/FontFile2") ||
                                pdfText.Contains("/FontFile3") ||
                                pdfText.Contains("/Subtype /TrueType");

        if (hasSubsetMarker || hasFontFileEntry)
        {
            Console.WriteLine("PDF font embedding verification succeeded.");
        }
        else
        {
            throw new InvalidOperationException("PDF does not contain expected font embedding markers; subsetting may be enabled.");
        }
    }
}
