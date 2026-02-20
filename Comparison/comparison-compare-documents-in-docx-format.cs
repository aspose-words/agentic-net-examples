using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and the edited documents (DOCX format).
        Document original = new Document("Original.docx");
        Document edited = new Document("Edited.docx");

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Example: compare moves, ignore formatting, etc.
            CompareMoves = true,
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            // Use the edited document as the target for comparison.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to the original document.
        original.Compare(edited, "Comparer", DateTime.Now, compareOptions);

        // Save the result with revisions to a new DOCX file.
        original.Save("ComparedResult.docx");
    }
}
