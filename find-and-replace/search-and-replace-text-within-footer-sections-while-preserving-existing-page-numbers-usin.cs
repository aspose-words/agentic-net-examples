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

        // Add enough content to generate multiple pages.
        for (int i = 0; i < 3; i++)
        {
            builder.Writeln($"This is page {i + 1}.");
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Create a primary footer with placeholder text and page fields.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Confidential - Page ");
        builder.InsertField("PAGE", null);
        builder.Write(" of ");
        builder.InsertField("NUMPAGES", null);
        builder.Writeln(); // End the paragraph.

        // Save the original document (optional, for inspection).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "OriginalFooter.docx");
        doc.Save(originalPath);

        // Prepare find-and-replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Locate the primary footer.
        HeaderFooter footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        if (footer == null)
            throw new InvalidOperationException("Footer not found.");

        // Replace the placeholder text while leaving the page fields untouched.
        int replacedCount = footer.Range.Replace("Confidential", "Public", options);

        // Validate that a replacement actually occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No occurrences of the target text were found in the footer.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FooterReplaceOutput.docx");
        doc.Save(outputPath);

        // Simple confirmation (no interactive prompts).
        Console.WriteLine($"Footer text replaced in {replacedCount} location(s). Output saved to: {outputPath}");
    }
}
