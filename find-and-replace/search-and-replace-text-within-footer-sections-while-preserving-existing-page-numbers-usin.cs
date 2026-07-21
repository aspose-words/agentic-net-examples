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

        // Add enough content to generate multiple pages.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i}.");
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Move to the primary footer of the first section and write placeholder text.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Confidential - Page ");
        builder.InsertField("PAGE", "?");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Access the primary footer.
        HeaderFooter footer = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];

        // Configure find-and-replace options (case‑insensitive, whole‑word not required).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Replace the placeholder text while leaving the PAGE field untouched.
        int replacedCount = footer.Range.Replace("Confidential", "Public", options);

        // Verify that a replacement actually occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement in the footer.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
