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

        // Enable a different header/footer for the first page of the section.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Create the header that will appear only on the first page.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("First page header");

        // Create the primary header that will appear on all other pages.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Primary header for other pages");

        // Return to the main body of the document and add three pages.
        builder.MoveToSection(0);
        builder.Writeln("Page 1 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content");

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstPageHeader.docx");
        doc.Save(outputPath);
    }
}
