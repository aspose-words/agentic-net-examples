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
        // Create a sample DOCX document with text using a TrueType font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("Hello World! This document is rendered to PDF with full font embedding.");

        // Define PDF save options to embed full fonts and disable subsetting.
        string pdfPath = "output.pdf";
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed the complete font (no subsetting).
            EmbedFullFonts = true
            // FontEmbeddingMode property is not required for full embedding in this version.
        };

        // Render the DOCX to PDF.
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new Exception("PDF file was not created.");

        // Load the PDF content as text for inspection.
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfContent = Encoding.ASCII.GetString(pdfBytes);

        // Check for embedded font markers in the PDF.
        bool hasEmbeddedFont = pdfContent.Contains("/FontFile") ||
                               pdfContent.Contains("/FontFile2") ||
                               pdfContent.Contains("/FontFile3") ||
                               pdfContent.Contains("/Subtype /TrueType");

        if (!hasEmbeddedFont)
            throw new Exception("Embedded font markers were not found in the PDF.");

        // Verify that no subset font name pattern (e.g., ABCDEF+FontName) exists.
        Regex subsetRegex = new Regex(@"[A-Z]{6}\+");
        bool hasSubsetPrefix = subsetRegex.IsMatch(pdfContent);

        if (hasSubsetPrefix)
            throw new Exception("Subset font prefix detected; subsetting is not disabled.");

        // Success message.
        Console.WriteLine("PDF rendered with full font embedding and subsetting disabled.");
    }
}
