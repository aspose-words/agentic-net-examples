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

        // Add three pages of content to demonstrate the different headers.
        builder.MoveToSection(0);
        builder.Writeln("Content of page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 3");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FirstPageHeader.docx");
        doc.Save(outputPath);
    }
}
