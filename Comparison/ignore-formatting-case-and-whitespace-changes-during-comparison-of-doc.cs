using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocuments
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Configure comparison options to ignore formatting, case changes and whitespace differences.
        // IgnoreFormatting skips any style/format changes.
        // IgnoreCaseChanges makes the comparison case‑insensitive.
        // Whitespace changes are treated as formatting differences, so they are ignored as well.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting   = true,
            IgnoreCaseChanges  = true,
            // Additional flags can be set to ignore other element types if desired.
            CompareMoves       = false,
            IgnoreComments     = false,
            IgnoreTables       = false,
            IgnoreFields       = false,
            IgnoreFootnotes    = false,
            IgnoreTextboxes    = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. The revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the result document that contains the tracked changes.
        docOriginal.Save("ComparisonResult.docx");
    }
}
