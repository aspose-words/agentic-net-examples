using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and edited DOCX documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure custom comparison options.
            CompareOptions options = new CompareOptions
            {
                // Track moved text as revisions.
                CompareMoves = true,
                // Ignore formatting differences.
                IgnoreFormatting = true,
                // Perform case‑insensitive comparison.
                IgnoreCaseChanges = true,
                // Include comments in the comparison.
                IgnoreComments = false,
                // Include tables in the comparison.
                IgnoreTables = false,
                // Ignore differences in fields.
                IgnoreFields = true,
                // Include footnotes/endnotes in the comparison.
                IgnoreFootnotes = false,
                // Include textboxes in the comparison.
                IgnoreTextboxes = false,
                // Include headers and footers in the comparison.
                IgnoreHeadersAndFooters = false,
                // Track changes at the word level.
                Granularity = Granularity.WordLevel,
                // Use the edited document as the target for comparison.
                Target = ComparisonTargetType.New
            };

            // Perform the document comparison.
            docOriginal.Compare(docEdited, "JD", DateTime.Now, options);
        }

        // Save the original document, which now contains revision marks.
        docOriginal.Save("ComparisonResult.docx");
    }
}
