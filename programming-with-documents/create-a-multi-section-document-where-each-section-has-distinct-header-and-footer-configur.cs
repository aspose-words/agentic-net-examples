using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Section 1 ----------
        // Header for Section 1.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 1");

        // Footer for Section 1.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 1");

        // Return to the body of Section 1 and add some content.
        builder.MoveToSection(0);
        builder.Writeln("Content of Section 1.");

        // Insert a section break to start Section 2 on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 2 ----------
        // Header for Section 2.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 2");

        // Footer for Section 2.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 2");

        // Return to the body of Section 2 and add some content.
        builder.MoveToSection(1);
        builder.Writeln("Content of Section 2.");

        // Insert a section break to start Section 3 on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 3 ----------
        // Enable different first page header/footer for this section.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // First page header for Section 3.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("First Page Header - Section 3");

        // First page footer for Section 3.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
        builder.Write("First Page Footer - Section 3");

        // Primary (other pages) header for Section 3.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Primary Header - Section 3");

        // Primary (other pages) footer for Section 3.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Primary Footer - Section 3");

        // Return to the body of Section 3 and add some content.
        builder.MoveToSection(2);
        builder.Writeln("Content of Section 3.");

        // Prepare output directory.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "MultiSectionHeadersFooters.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
