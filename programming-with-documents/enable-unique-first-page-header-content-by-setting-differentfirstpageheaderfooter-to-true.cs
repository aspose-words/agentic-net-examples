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

        // Enable different first page header/footer for the first section.
        Section section = doc.FirstSection;
        section.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Add content to the first page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("First Page Header");

        // Add content to the primary (default) header for the rest of the pages.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Primary Header");

        // Add some body text to generate multiple pages.
        builder.MoveToDocumentEnd();
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");

        // Save the document to disk.
        string outputPath = "FirstPageHeader.docx";
        doc.Save(outputPath);
    }
}
