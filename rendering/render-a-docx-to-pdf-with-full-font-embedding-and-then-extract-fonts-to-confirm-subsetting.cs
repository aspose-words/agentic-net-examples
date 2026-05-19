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
        // Define paths for the sample DOCX and the resulting PDF.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docxPath = Path.Combine(outputDir, "Sample.docx");
        string pdfPath = Path.Combine(outputDir, "Sample_FullFonts.pdf");

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX that uses a TrueType font.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";               // A TrueType font available on most systems.
        builder.Writeln("This text will be rendered with full font embedding.");
        doc.Save(docxPath);                         // Save the source DOCX.

        // -----------------------------------------------------------------
        // 2. Render the DOCX to PDF with full font embedding (no subsetting).
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true   // Disable subsetting; embed the complete font.
        };
        doc.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 3. Verify that the PDF contains embedded font markers.
        //    We look for /FontFile, /FontFile2, /FontFile3 or a TrueType subtype.
        //    The absence of a subset prefix (e.g., ABCDEF+) indicates subsetting is disabled.
        // -----------------------------------------------------------------
        string pdfContent = File.ReadAllText(pdfPath, Encoding.ASCII);

        bool hasFontFileMarker = pdfContent.Contains("/FontFile") ||
                                 pdfContent.Contains("/FontFile2") ||
                                 pdfContent.Contains("/FontFile3");

        bool hasTrueTypeSubtype = pdfContent.Contains("/Subtype /TrueType");

        bool hasSubsetPrefix = Regex.IsMatch(pdfContent, @"[A-Z]{6}\+");

        if (!hasFontFileMarker && !hasTrueTypeSubtype)
        {
            throw new InvalidOperationException("The PDF does not contain any embedded font markers.");
        }

        if (hasSubsetPrefix)
        {
            throw new InvalidOperationException("The PDF appears to contain subsetted fonts, but full embedding was requested.");
        }

        Console.WriteLine("PDF generated with full font embedding successfully.");
        Console.WriteLine($"PDF path: {pdfPath}");
    }
}
