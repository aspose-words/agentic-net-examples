using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the left margin for the main body (e.g., 1 inch).
        builder.PageSetup.LeftMargin = ConvertUtil.InchToPoint(1.0);

        // Enable a distinct header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move the cursor to the first‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Insert header text.
        builder.Writeln("First Page Header");

        // Align the header's left indent with the body left margin.
        builder.CurrentParagraph.ParagraphFormat.LeftIndent = builder.PageSetup.LeftMargin;

        // Return to the main section body.
        builder.MoveToSection(0);
        builder.Writeln("Body text starts here.");

        // Save the document to a file.
        doc.Save("FirstPageHeaderAligned.docx");
    }
}
