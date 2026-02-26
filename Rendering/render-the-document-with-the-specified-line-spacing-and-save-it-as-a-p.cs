using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderWithLineSpacing
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and configure line spacing.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the line spacing rule to "Multiple" (i.e., a multiple of the default line height).
        // Here we use 1.5 lines, which corresponds to 18 points (12 points * 1.5).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // 1.5 * 12 = 18 points

        // Add some sample paragraphs to demonstrate the spacing.
        builder.Writeln("First paragraph with 1.5 line spacing.");
        builder.Writeln("Second paragraph with 1.5 line spacing.");
        builder.Writeln("Third paragraph with 1.5 line spacing.");

        // Prepare PDF save options (default options are sufficient for this task).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as a PDF file using the specified options.
        doc.Save("RenderedDocument.pdf", pdfOptions);
    }
}
