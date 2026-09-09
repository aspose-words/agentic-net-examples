using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Section 1 ----------
        // Create header and footer for the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 1");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 1");

        // Add some body content to section 1.
        builder.MoveToSection(0);
        builder.Writeln("Content of Section 1");

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 2 ----------
        // Create header and footer for the second section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 2");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 2");

        // Add body content to section 2.
        builder.MoveToSection(1);
        builder.Writeln("Content of Section 2");

        // Insert a section break to start Section 3.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 3 ----------
        // Create header and footer for the third section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 3");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 3");

        // Add body content to section 3.
        builder.MoveToSection(2);
        builder.Writeln("Content of Section 3");

        // Save the document to the current directory.
        doc.Save("MultiSectionHeadersFooters.docx");
    }
}
