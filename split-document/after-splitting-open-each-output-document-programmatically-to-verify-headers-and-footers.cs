using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document with two sections and headers/footers.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Enable different headers/footers for first page and odd/even pages.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // ----- First section -----
        // Header – first page
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("First page header");

        // Header – even pages
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Write("Even page header");

        // Header – primary (odd) pages
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Primary page header");

        // Footer – primary (odd) pages
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Primary page footer");

        // Body content for first section
        builder.MoveToDocumentEnd();
        builder.Writeln("Content of the first section.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the first section.");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ----- Second section -----
        // Header – primary (odd) pages for second section
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Second section header");

        // Footer – primary (odd) pages for second section
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Second section footer");

        // Body content for second section
        builder.MoveToDocumentEnd();
        builder.Writeln("Content of the second section.");

        // -----------------------------------------------------------------
        // 2. Split the document by sections, preserving headers and footers.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document splitDoc = new Document();
            splitDoc.EnsureMinimum();

            // Import the current section into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(sourceDoc.Sections[i], true);

            // Replace the default empty section with the imported one.
            splitDoc.Sections.Clear();
            splitDoc.Sections.Add(importedSection);

            // Save the split part.
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            splitDoc.Save(partPath);
        }

        // -----------------------------------------------------------------
        // 3. Verify that each split document contains the expected headers/footers.
        // -----------------------------------------------------------------
        Console.WriteLine("Verification of split documents:");
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            Document partDoc = new Document(partPath);

            // Retrieve primary header and footer (if they exist).
            HeaderFooter header = partDoc.FirstSection?.HeadersFooters[HeaderFooterType.HeaderPrimary];
            HeaderFooter footer = partDoc.FirstSection?.HeadersFooters[HeaderFooterType.FooterPrimary];

            string headerText = header?.GetText().Trim() ?? "(no header)";
            string footerText = footer?.GetText().Trim() ?? "(no footer)";

            Console.WriteLine($"  Part {i + 1}:");
            Console.WriteLine($"    Header: {headerText}");
            Console.WriteLine($"    Footer: {footerText}");
        }
    }
}
