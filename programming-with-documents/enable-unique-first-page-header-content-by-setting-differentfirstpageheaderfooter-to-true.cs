using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a different header/footer for the first page of the section.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // ----- First‑page header -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("Header for the first page");

        // ----- Primary header (used on all other pages) -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header for subsequent pages");

        // Return to the main body of the first section.
        builder.MoveToSection(0);

        // Add three pages to demonstrate the different headers.
        builder.Writeln("Content of page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 3");

        // Save the document to the local file system.
        doc.Save("FirstPageHeader.docx");
    }
}
