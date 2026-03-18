using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a different header/footer for the first page of each section.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // ----- First‑page header -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("Header for the first page");

        // ----- Primary header (used on all other pages) -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header for subsequent pages");

        // Add some content and page breaks to see the headers in action.
        builder.MoveToSection(0);
        builder.Writeln("Content of page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 3");

        // Save the document to disk.
        const string outputPath = "FirstPageHeader.docx";
        doc.Save(outputPath);
    }
}
