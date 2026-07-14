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
        // Define folders for the sample document and the resulting PDF.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the generated PDF file.
        string pdfPath = Path.Combine(outputDir, "SampleSubsetFonts.pdf");

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document that uses a TrueType font.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a common TrueType font. The exact font does not matter; Aspose.Words
        // will embed the subset of glyphs that are actually used.
        builder.Font.Name = "Arial";
        builder.Font.Size = 24;
        builder.Writeln("Hello, Aspose.Words!");
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // -----------------------------------------------------------------
        // 2. Configure PDF save options to embed only the used glyphs.
        //    Subsetting is the default behavior (EmbedFullFonts = false).
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Explicitly set to false for clarity – only used glyphs will be embedded.
            EmbedFullFonts = false,
            // Ensure that core fonts are not substituted, so the TrueType font is embedded.
            UseCoreFonts = false
        };

        // Save the document as PDF with the specified options.
        doc.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 3. Verify that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // -----------------------------------------------------------------
        // 4. Inspect the PDF content for markers that indicate font subsetting.
        //    Subset fonts in PDFs are named with a six‑letter prefix followed by '+'.
        //    Additionally, the presence of '/FontFile' or '/FontFile2' confirms embedding.
        // -----------------------------------------------------------------
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfText = Encoding.ASCII.GetString(pdfBytes);

        bool hasSubsetFontName = Regex.IsMatch(pdfText, @"[A-Z]{6}\+");
        bool hasEmbeddedFontMarker = pdfText.Contains("/FontFile") || pdfText.Contains("/FontFile2");

        if (!hasSubsetFontName && !hasEmbeddedFontMarker)
            throw new InvalidOperationException("The PDF does not contain embedded subset font markers.");

        // -----------------------------------------------------------------
        // 5. Output a simple confirmation (no interactive input required).
        // -----------------------------------------------------------------
        Console.WriteLine("PDF generated successfully at:");
        Console.WriteLine(pdfPath);
        Console.WriteLine("Embedded subset font markers detected: " + (hasSubsetFontName ? "Yes" : "No"));
    }
}
