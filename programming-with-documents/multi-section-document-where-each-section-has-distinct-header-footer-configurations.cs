using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class MultiSectionHeadersFooters
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Section 1 ----------
        // Enable different headers/footers for first, even and odd pages.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // First page header/footer.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("Section 1 – First Page Header");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
        builder.Write("Section 1 – First Page Footer");

        // Even page header/footer.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Write("Section 1 – Even Page Header");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
        builder.Write("Section 1 – Even Page Footer");

        // Primary (odd) page header/footer.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Section 1 – Primary Header");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Section 1 – Primary Footer");

        // Add some body content for the first section.
        builder.MoveToSection(0);
        builder.Writeln("Content of Section 1 – Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of Section 1 – Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of Section 1 – Page 3");

        // ---------- Section 2 ----------
        // Insert a section break to start a new section on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Unlink the new section's headers/footers from the previous section.
        doc.Sections[1].HeadersFooters.LinkToPrevious(false);

        // Configure distinct headers/footers for Section 2.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Section 2 – Primary Header");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Section 2 – Primary Footer");

        // (Optional) Different first page header/footer for Section 2.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("Section 2 – First Page Header");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
        builder.Write("Section 2 – First Page Footer");

        // Add body content for the second section.
        builder.MoveToSection(1);
        builder.Writeln("Content of Section 2 – Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of Section 2 – Page 2");

        // ---------- Section 3 ----------
        // Insert another section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Unlink Section 3 headers/footers from Section 2.
        doc.Sections[2].HeadersFooters.LinkToPrevious(false);

        // Distinct headers/footers for Section 3.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Section 3 – Primary Header");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Section 3 – Primary Footer");

        // No different first page header/footer for this section (uses primary header/footer).

        // Add body content for the third section.
        builder.MoveToSection(2);
        builder.Writeln("Content of Section 3 – Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of Section 3 – Page 2");

        // Save the document.
        doc.Save("MultiSectionHeadersFooters.docx");
    }
}
