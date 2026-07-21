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

        // -----------------------------------------------------------------
        // Section 1
        // -----------------------------------------------------------------
        // Add some body text.
        builder.Writeln("Section 1 – Body content.");

        // Create a primary header for Section 1.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header – Section 1");

        // Create a primary footer for Section 1.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer – Section 1");

        // Return the cursor to the body of Section 1.
        builder.MoveToSection(0);
        builder.Writeln("More text in Section 1.");

        // Insert a section break to start Section 2 on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // -----------------------------------------------------------------
        // Section 2
        // -----------------------------------------------------------------
        // Add body text for Section 2.
        builder.Writeln("Section 2 – Body content.");

        // Header for Section 2.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header – Section 2");

        // Footer for Section 2.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer – Section 2");

        // Return to the body of Section 2.
        builder.MoveToSection(1);
        builder.Writeln("More text in Section 2.");

        // Insert a section break to start Section 3 on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // -----------------------------------------------------------------
        // Section 3
        // -----------------------------------------------------------------
        // Add body text for Section 3.
        builder.Writeln("Section 3 – Body content.");

        // Header for Section 3.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header – Section 3");

        // Footer for Section 3.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer – Section 3");

        // Return to the body of Section 3.
        builder.MoveToSection(2);
        builder.Writeln("More text in Section 3.");

        // -----------------------------------------------------------------
        // Save the document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MultiSectionHeadersFooters.docx");
        doc.Save(outputPath);
    }
}
