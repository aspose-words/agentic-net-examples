using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocuments
{
    static void Main()
    {
        // Load the original and the edited documents (DOCX format).
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Configure comparison options to ignore all possible changes.
        CompareOptions compareOptions = new CompareOptions
        {
            // Do not generate move revisions.
            CompareMoves = false,
            // Ignore formatting differences.
            IgnoreFormatting = true,
            // Ignore case changes.
            IgnoreCaseChanges = true,
            // Ignore comments, tables, fields, footnotes, textboxes, and headers/footers.
            IgnoreComments = true,
            IgnoreTables = true,
            IgnoreFields = true,
            IgnoreFootnotes = true,
            IgnoreTextboxes = true,
            IgnoreHeadersAndFooters = true,
            // Use the edited document as the target for comparison.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added only for changes not ignored above.
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Save the comparison result to a new DOCX file.
        docOriginal.Save("ComparisonResult.docx");
    }
}
