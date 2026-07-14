using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // File names (created in the current working directory).
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create a sample document with several headers.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable distinct first‑page and odd/even headers.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Primary header (odd pages) – contains the keyword.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Company Report – ReplaceMe");

        // First‑page header – does NOT contain the keyword.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("Company Report – First Page");

        // Even‑page header – contains the keyword.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Writeln("Company Report – ReplaceMe");

        // Return to the main story before inserting a section break.
        builder.MoveToDocumentEnd();
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section – same header layout.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Summary – ReplaceMe");

        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("Summary – First Page");

        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Writeln("Summary – ReplaceMe");

        // Save the sample document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document and replace the keyword only in headers.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        const string keyword = "ReplaceMe";
        const string replacement = "Replaced";

        int totalReplacements = 0;
        FindReplaceOptions options = new FindReplaceOptions();

        foreach (Section section in loadedDoc.Sections)
        {
            HeaderFooterCollection headers = section.HeadersFooters;
            foreach (HeaderFooter header in headers)
            {
                // Process only headers that actually contain the keyword.
                if (header.Range.Text.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    int replaced = header.Range.Replace(keyword, replacement, options);
                    totalReplacements += replaced;
                }
            }
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No header replacements were performed.");

        // Save the modified document.
        loadedDoc.Save(outputPath);
    }
}
