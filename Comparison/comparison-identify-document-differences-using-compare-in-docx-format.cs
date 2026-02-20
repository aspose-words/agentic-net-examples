using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original document.
        Document docOriginal = new Document("Original.docx");

        // Load the edited document that we want to compare against.
        Document docEdited = new Document("Edited.docx");

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the word level.
            Granularity = Granularity.WordLevel,

            // Do not ignore any element types – track all differences.
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,

            // Use the original document as the base for comparison.
            Target = ComparisonTargetType.Current,

            // Do not treat moved text as a separate change.
            CompareMoves = false
        };

        // Perform the comparison. Revisions (tracked changes) will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the document that now contains revision marks.
        docOriginal.Save("ComparisonResult.docx");
    }
}
