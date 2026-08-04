using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Section 1 ----------
        // Header for Section 1
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 1");

        // Footer for Section 1
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 1");

        // Return to the body of Section 1 and add some content.
        builder.MoveToSection(0);
        builder.Writeln("Content of Section 1");

        // Insert a section break (new page) to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 2 ----------
        // Header for Section 2
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 2");

        // Footer for Section 2
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 2");

        // Return to the body of Section 2 and add content.
        builder.MoveToSection(1);
        builder.Writeln("Content of Section 2");

        // Insert another section break to start Section 3.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 3 ----------
        // Header for Section 3
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 3");

        // Footer for Section 3
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 3");

        // Return to the body of Section 3 and add content.
        builder.MoveToSection(2);
        builder.Writeln("Content of Section 3");

        // Save the document to the local file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MultiSectionHeadersFooters.docx");
        doc.Save(outputPath);
    }
}
