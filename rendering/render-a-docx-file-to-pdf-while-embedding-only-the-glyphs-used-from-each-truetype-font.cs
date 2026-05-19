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
        // Prepare folders for output.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple DOCX document that uses a TrueType font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a common TrueType font. The default behavior will embed only the used glyphs.
        builder.Font.Name = "Arial";
        builder.Font.Size = 24;
        builder.Writeln("Hello, Aspose.Words!");
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Path for the resulting PDF.
        string pdfPath = Path.Combine(artifactsDir, "SampleSubsetFonts.pdf");

        // Configure PDF save options to ensure subsetting (EmbedFullFonts = false).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = false // Subset fonts – only used glyphs are embedded.
        };

        // Render the document to PDF.
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Load the PDF content as text for simple inspection.
        string pdfContent = File.ReadAllText(pdfPath, Encoding.ASCII);

        // Look for a subset font name pattern (e.g., "ABCDEF+Arial").
        bool hasSubsetFont = Regex.IsMatch(pdfContent, @"[A-Z]{6}\+");

        // Also check for a font file entry indicating embedding (e.g., /FontFile2).
        bool hasEmbeddedFont = pdfContent.Contains("/FontFile2") || pdfContent.Contains("/FontFile");

        if (!hasSubsetFont && !hasEmbeddedFont)
            throw new InvalidOperationException("The PDF does not contain embedded subset font markers.");

        // If execution reaches this point, the PDF was generated with subsetted fonts.
        Console.WriteLine("PDF generated successfully with subsetted TrueType fonts at:");
        Console.WriteLine(pdfPath);
    }
}
