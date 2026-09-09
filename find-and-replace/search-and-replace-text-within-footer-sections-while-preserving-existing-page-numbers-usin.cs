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

        // Ensure the document has at least one section.
        doc.EnsureMinimum();

        // Add a primary footer to the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Write some placeholder text in the footer.
        builder.Write("Company XYZ - Confidential ");

        // Insert a page number field; this will be preserved during replacement.
        builder.InsertField("PAGE", "?");

        // Return to the main body for any further content (optional).
        builder.MoveToDocumentEnd();

        // Access the primary footer.
        HeaderFooter footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];

        // Set up find-and-replace options (case‑insensitive, replace whole words not required).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Replace the placeholder company name while leaving the page number untouched.
        int replacedCount = footer.Range.Replace("Company XYZ", "Acme Corp", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement in the footer.");

        // Define output path relative to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FooterReplaced.docx");

        // Save the modified document.
        doc.Save(outputPath);
    }
}
