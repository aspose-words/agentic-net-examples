using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocsWithRevisions
{
    static void Main()
    {
        // Load the original and the edited documents (DOCX format).
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Set up comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Do not treat moved text as a separate change.
            CompareMoves = false,
            // Track changes at the word level (use Granularity.CharLevel for character-level tracking).
            Granularity = Granularity.WordLevel,
            // Do not ignore any element types – all differences will generate revisions.
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            // Use the edited document as the target (base) for comparison.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, compareOptions);

        // Iterate over revision groups to display summary information.
        foreach (RevisionGroup group in docOriginal.Revisions.Groups)
        {
            Console.WriteLine($"Author: {group.Author}, Type: {group.RevisionType}, Text: {group.Text}");
        }

        // Save the resulting document (it contains the tracked changes).
        docOriginal.Save("ComparedResult.docx");
    }
}
