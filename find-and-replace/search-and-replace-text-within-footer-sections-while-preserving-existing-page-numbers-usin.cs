using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three pages of sample content.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is the body text of page {i}.");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Move to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        // Write some static text and insert a PAGE field for page numbers.
        builder.Write("Confidential - CompanyName - Page ");
        builder.InsertField("PAGE", "");

        // Ensure the same footer appears on all pages (same section, same footer).
        // No additional code needed because we edited the primary footer of the section.

        // Prepare find-and-replace options (case‑insensitive, whole‑word not required).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Perform the replacement only in the footer range.
        HeaderFooter footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        int replacedCount = footer.Range.Replace("CompanyName", "Acme Corp", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement in the footer.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
