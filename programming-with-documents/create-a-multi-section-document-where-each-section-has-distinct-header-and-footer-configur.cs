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
        // Primary header and footer for the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 1");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 1");

        // Add some body content.
        builder.MoveToSection(0);
        builder.Writeln("Content of Section 1");

        // Start a new section on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 2 ----------
        // Enable a different first page header/footer.
        builder.MoveToSection(1);
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // First page header/footer.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("First Page Header - Section 2");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
        builder.Write("First Page Footer - Section 2");

        // Primary header/footer for the rest of the pages in this section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 2");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 2");

        // Add body content.
        builder.MoveToSection(1);
        builder.Writeln("Content of Section 2");

        // Start another new section on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 3 ----------
        // Enable different headers/footers for odd and even pages.
        builder.MoveToSection(2);
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Even page header/footer.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Write("Even Header - Section 3");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
        builder.Write("Even Footer - Section 3");

        // Primary (odd) page header/footer.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Primary Header - Section 3");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Primary Footer - Section 3");

        // Add body content.
        builder.MoveToSection(2);
        builder.Writeln("Content of Section 3");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "MultiSectionHeadersFooters.docx");
        doc.Save(outputPath);
    }
}
