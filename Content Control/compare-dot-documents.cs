using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original document.
        Document docOriginal = new Document("Original.docx");

        // Load the edited document to compare against.
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");
        }

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the word level (default). Change to CharLevel for character granularity.
            Granularity = Granularity.WordLevel,

            // Do not ignore any element types; all differences will be captured.
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,

            // Use the edited document as the target for comparison (equivalent to Word's "Show changes in New").
            Target = ComparisonTargetType.New,

            // Do not generate move revisions.
            CompareMoves = false
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Optionally, accept all revisions to transform the original into the edited version.
        // docOriginal.Revisions.AcceptAll();

        // Save the resulting document with tracked changes.
        docOriginal.Save("ComparisonResult.docx");
    }
}
