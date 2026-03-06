using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure line spacing: 1.5 times the default line height (12 points).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // 1.5 * 12 points

        // Add sample paragraphs to demonstrate the spacing.
        builder.Writeln("First line with custom line spacing.");
        builder.Writeln("Second line with the same custom line spacing.");

        // Create PDF save options (default settings are sufficient for this task).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as a PDF file.
        doc.Save("RenderedDocument.pdf", pdfOptions);
    }
}
