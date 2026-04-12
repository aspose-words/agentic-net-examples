using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different headers for first page and even pages.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Primary header (contains the keyword "Company").
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Company Confidential");

        // First page header (does NOT contain the keyword).
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("First Page Header");

        // Even page header (contains the keyword "Company").
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Writeln("Company Internal");

        // Add a body paragraph so the document has content.
        builder.MoveToDocumentEnd();
        builder.Writeln("Body content goes here.");

        // Define the keyword to search for and its replacement.
        const string keyword = "Company";
        const string replacement = "Acme Corp";

        int totalReplacements = 0;

        // Iterate through all sections and their headers.
        foreach (Section section in doc.Sections)
        {
            HeaderFooterCollection headers = section.HeadersFooters;
            foreach (HeaderFooter header in headers)
            {
                // Process only header types (skip footers).
                if (header.HeaderFooterType != HeaderFooterType.FooterPrimary &&
                    header.HeaderFooterType != HeaderFooterType.FooterFirst &&
                    header.HeaderFooterType != HeaderFooterType.FooterEven)
                {
                    // If the header contains the keyword, replace it.
                    if (header.Range.Text.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    {
                        int count = header.Range.Replace(keyword, replacement);
                        totalReplacements += count;
                    }
                }
            }
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No header replacements were performed.");

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ModifiedHeaders.docx");
        doc.Save(outputPath);
    }
}
