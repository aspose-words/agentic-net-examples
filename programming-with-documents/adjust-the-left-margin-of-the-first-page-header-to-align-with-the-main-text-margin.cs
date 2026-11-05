using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the left margin for the main body (e.g., 72 points = 1 inch).
        Section section = doc.FirstSection;
        section.PageSetup.LeftMargin = 72; // 1 inch

        // Write some body text.
        builder.Writeln("This is the main body text. It uses the left margin defined for the page.");

        // Enable a different first‑page header.
        section.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Insert the first‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        // Align the header's left margin with the main text margin.
        builder.ParagraphFormat.LeftIndent = section.PageSetup.LeftMargin;
        builder.Writeln("First‑page header aligned with main text margin.");

        // Return to the main document body.
        builder.MoveToDocumentEnd();

        // Save the document.
        string outputPath = "FirstPageHeaderMargin.docx";
        doc.Save(outputPath);
    }
}
