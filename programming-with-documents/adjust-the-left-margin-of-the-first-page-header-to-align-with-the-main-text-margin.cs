using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the main text left margin (e.g., 1 inch).
        doc.FirstSection.PageSetup.LeftMargin = ConvertUtil.InchToPoint(1.0);

        // Add some body text.
        builder.Writeln("This is the main body text.");

        // Enable a different first‑page header.
        doc.FirstSection.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Create the first‑page header.
        HeaderFooter firstHeader = new HeaderFooter(doc, HeaderFooterType.HeaderFirst);
        doc.FirstSection.HeadersFooters.Add(firstHeader);

        // Add a paragraph to the first‑page header.
        Paragraph headerParagraph = new Paragraph(doc);
        Run run = new Run(doc, "First page header text");
        headerParagraph.AppendChild(run);
        firstHeader.AppendChild(headerParagraph);

        // Align the header's left margin with the main text margin.
        headerParagraph.ParagraphFormat.LeftIndent = doc.FirstSection.PageSetup.LeftMargin;

        // Save the document.
        string outputPath = "AdjustedHeader.docx";
        doc.Save(outputPath);
    }
}
