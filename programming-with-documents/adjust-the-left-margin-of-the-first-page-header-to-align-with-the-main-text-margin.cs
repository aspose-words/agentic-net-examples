using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set page margins (1 inch = 72 points).
        builder.PageSetup.LeftMargin = 72;
        builder.PageSetup.RightMargin = 72;
        builder.PageSetup.TopMargin = 72;
        builder.PageSetup.BottomMargin = 72;

        // Insert body text.
        builder.Writeln("This is the main body text. It should respect the page margins.");

        // Move to the first-page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        // Insert header text.
        builder.Writeln("First Page Header");

        // Align the header's left margin with the main text left margin.
        double leftMargin = doc.FirstSection.PageSetup.LeftMargin;
        builder.ParagraphFormat.LeftIndent = leftMargin;

        // Save the document.
        string outputPath = "AdjustedHeader.docx";
        doc.Save(outputPath);
    }
}
