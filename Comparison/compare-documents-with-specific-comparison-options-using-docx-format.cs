using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocuments
{
    static void Main()
    {
        // Load the two DOCX documents to be compared.
        Document original = new Document("Original.docx");
        Document revised = new Document("Revised.docx");

        // Configure comparison options.
        CompareOptions options = new CompareOptions
        {
            // Do not track moved content.
            CompareMoves = false,
            // Track changes at the word level.
            Granularity = Granularity.WordLevel,
            // Ignore case differences.
            IgnoreCaseChanges = true,
            // Ignore comments, formatting, footnotes, headers/footers, tables, textboxes, and fields.
            IgnoreComments = true,
            IgnoreFormatting = true,
            IgnoreFootnotes = true,
            IgnoreHeadersAndFooters = true,
            IgnoreTables = true,
            IgnoreTextboxes = true,
            IgnoreFields = true,
            // Use the revised document as the base for comparison.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions are added to the 'original' document.
        original.Compare(revised, "Reviewer", DateTime.Now, options);

        // Save the document that now contains the revision marks.
        original.Save("ComparisonResult.docx");
    }
}
