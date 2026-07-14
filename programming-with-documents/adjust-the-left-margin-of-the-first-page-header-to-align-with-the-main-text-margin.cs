using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a custom left margin for the main body (1 inch = 72 points).
        double leftMargin = ConvertUtil.InchToPoint(1.0);
        builder.PageSetup.LeftMargin = leftMargin;

        // Enable a different header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move the cursor to the first‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Write header text.
        builder.Writeln("First Page Header");

        // Align the header text with the main body left margin by setting the paragraph's left indent.
        builder.CurrentParagraph.ParagraphFormat.LeftIndent = leftMargin;

        // Return to the main document body.
        builder.MoveToSection(0);
        builder.Writeln("This is the main document body. The first‑page header aligns with the left margin.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FirstPageHeaderMargin.docx");
        doc.Save(outputPath);
    }
}
