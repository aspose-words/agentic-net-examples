using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class FooterFindReplaceDemo
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some body content.
        builder.Writeln("This is the main document body.");
        builder.Writeln("It spans multiple pages to demonstrate page numbers in footers.");

        // Insert a primary footer with placeholder text and a PAGE field.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Confidential - Draft - Page ");
        builder.InsertField("PAGE", "?");

        // Return to the main body to add more pages.
        builder.MoveToDocumentEnd();
        for (int i = 0; i < 3; i++)
        {
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Additional content on page {i + 2}.");
        }

        // Save the original document (optional, demonstrates lifecycle usage).
        const string inputPath = "FooterOriginal.docx";
        doc.Save(inputPath);

        // Load the document back (simulating a typical load‑modify‑save scenario).
        Document loadedDoc = new Document(inputPath);

        // Access the primary footer of the first section.
        HeaderFooter footer = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];

        // Configure find‑replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Replace the word "Draft" with "Final" inside the footer.
        int replacedCount = footer.Range.Replace("Draft", "Final", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement in the footer.");

        // Save the modified document.
        const string outputPath = "FooterReplaced.docx";
        loadedDoc.Save(outputPath);
    }
}
