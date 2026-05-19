using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Access the builder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a custom left margin for the main body (1 inch = 72 points).
        builder.PageSetup.LeftMargin = ConvertUtil.InchToPoint(1.0);

        // Enable a different header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move to the first‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Write header text.
        builder.Write("First Page Header");

        // Align the header's left margin with the main text margin.
        // The left indent of the paragraph inside the header is set to the section's left margin.
        builder.ParagraphFormat.LeftIndent = doc.FirstSection.PageSetup.LeftMargin;

        // Return to the main document body.
        builder.MoveToSection(0);
        builder.Writeln("Body text that follows the header.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AdjustedHeaderMargin.docx");
        doc.Save(outputPath);
    }
}
