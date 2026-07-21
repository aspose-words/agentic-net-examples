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
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple DOCX document that uses a TrueType font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";               // Common TrueType font.
        builder.Writeln("Hello world! This text will be rendered with full font embedding.");

        // Configure PDF save options to embed the full font (no subsetting).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Save the document as PDF.
        string pdfPath = Path.Combine(outputDir, "FullFontEmbedding.pdf");
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Load the PDF content as ASCII text for simple inspection.
        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));

        // Check for font embedding markers (e.g., /FontFile, /FontFile2, /FontFile3).
        bool hasFontFileMarker = pdfContent.Contains("/FontFile") ||
                                 pdfContent.Contains("/FontFile2") ||
                                 pdfContent.Contains("/FontFile3");

        if (!hasFontFileMarker)
            throw new InvalidOperationException("The PDF does not contain any font embedding markers.");

        // Verify that subsetting is disabled by ensuring no subset-style font names are present.
        // Subset fonts are usually indicated by six uppercase letters followed by a '+' (e.g., ABCDEF+FontName).
        Regex subsetPattern = new Regex(@"[A-Z]{6}\+");
        if (subsetPattern.IsMatch(pdfContent))
            throw new InvalidOperationException("Subset-style font names were found in the PDF, indicating subsetting is enabled.");

        // If we reach this point, the PDF contains full font embedding without subsetting.
        Console.WriteLine("PDF generated successfully with full font embedding and without subsetting.");
    }
}
