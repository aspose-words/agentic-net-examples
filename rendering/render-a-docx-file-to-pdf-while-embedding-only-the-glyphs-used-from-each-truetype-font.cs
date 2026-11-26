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
        // Paths for the temporary DOCX and the resulting PDF
        const string docPath = "sample.docx";
        const string pdfPath = "output.pdf";

        // Create a sample DOCX that uses a TrueType font (e.g., Arial)
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This is a sample text using Arial font.");
        builder.Writeln("Additional characters: 1234567890!@#$%^&*()");
        doc.Save(docPath);

        // Configure PDF save options.
        // By default Aspose.Words embeds only the glyphs actually used (subset).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Render the DOCX to PDF with the specified options
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Simple verification: look for typical subset‑font markers in the PDF content
        string pdfContent = File.ReadAllText(pdfPath, Encoding.ASCII);
        bool hasSubsetFont = pdfContent.Contains("/FontFile2") ||
                             pdfContent.Contains("/FontFile3") ||
                             Regex.IsMatch(pdfContent, @"[A-Z]{6}\+");

        if (!hasSubsetFont)
            throw new InvalidOperationException("The PDF does not contain embedded subset fonts.");

        // (Optional) Clean up generated files
        // File.Delete(docPath);
        // File.Delete(pdfPath);
    }
}
