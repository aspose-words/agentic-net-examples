using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a custom left margin for the body text (1 inch).
        double leftMargin = ConvertUtil.InchToPoint(1.0);
        builder.PageSetup.LeftMargin = leftMargin;

        // Enable a different header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move to the first‑page header and add some text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("First page header");

        // Align the header's left margin with the main text margin.
        // The header is a separate story; retrieve its first paragraph and set its left indent.
        HeaderFooter firstHeader = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderFirst];
        if (firstHeader != null && firstHeader.FirstParagraph != null)
        {
            // LeftIndent is measured relative to the page margin.
            // Setting it to zero makes the header start at the same left edge as the body text.
            firstHeader.FirstParagraph.ParagraphFormat.LeftIndent = 0;
        }

        // Return to the main document body and add some content.
        builder.MoveToSection(0);
        builder.Writeln("Body text starts here.");

        // Save the document.
        doc.Save("AdjustedHeaderMargin.docx");
    }
}
