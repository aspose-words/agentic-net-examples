using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a distinct header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Define the left margin for the body (e.g., 1 inch).
        builder.PageSetup.LeftMargin = ConvertUtil.InchToPoint(1.0);

        // Move the cursor to the first‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Align the header text with the body left margin.
        builder.ParagraphFormat.LeftIndent = builder.PageSetup.LeftMargin;

        // Insert header content.
        builder.Write("First page header aligned with body margin");

        // Return to the main section body.
        builder.MoveToSection(0);
        builder.Writeln("Body text starts here.");

        // Save the document.
        doc.Save("FirstPageHeaderAligned.docx");
    }
}
