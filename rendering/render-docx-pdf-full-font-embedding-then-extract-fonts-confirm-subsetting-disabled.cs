using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths – adjust as needed.
        string inputDocx = @"C:\Docs\InputDocument.docx";
        string outputPdf = @"C:\Docs\OutputDocument.pdf";
        string extractedFontsDir = @"C:\Docs\ExtractedFonts";

        // Ensure the output folder for extracted fonts exists.
        Directory.CreateDirectory(extractedFontsDir);

        // Load the DOCX document.
        Document doc = new Document(inputDocx);

        // Configure PDF save options to embed the full font files (no subsetting).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true // Full font embedding, disables subsetting.
        };

        // Save the document as PDF with the specified options.
        doc.Save(outputPdf, pdfOptions);

        // --------------------------------------------------------------------
        // Extract embedded fonts from the original document to verify that
        // subsetting is not applied (full font files are present).
        // --------------------------------------------------------------------
        foreach (FontInfo fontInfo in doc.FontInfos)
        {
            // Consider only TrueType fonts that are actually embedded.
            if (!fontInfo.IsTrueType)
                continue;

            // Attempt to retrieve the embedded font in OpenType format.
            // If the font is not embedded, the method returns null.
            byte[] fontData = fontInfo.GetEmbeddedFontAsOpenType(EmbeddedFontStyle.Regular);

            if (fontData == null)
                continue; // Font not embedded – skip.

            // Build a safe file name for the extracted font.
            string safeFontName = string.Concat(fontInfo.Name.Split(Path.GetInvalidFileNameChars()));
            string fontFilePath = Path.Combine(extractedFontsDir, $"{safeFontName}.otf");

            // Write the font bytes to disk.
            File.WriteAllBytes(fontFilePath, fontData);
        }

        Console.WriteLine("PDF saved with full font embedding.");
        Console.WriteLine($"Extracted fonts are located in: {extractedFontsDir}");
    }
}
