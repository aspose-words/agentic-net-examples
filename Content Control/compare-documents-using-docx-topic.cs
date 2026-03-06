using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and the edited documents.
        Document original = new Document("Original.docx");
        Document edited   = new Document("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (original.Revisions.Count != 0 || edited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Set up comparison options (optional – customize as needed).
        CompareOptions options = new CompareOptions
        {
            // Example: track changes at the word level and ignore case changes.
            Granularity = Granularity.WordLevel,
            IgnoreCaseChanges = true,
            // Use the original document as the base for comparison.
            Target = ComparisonTargetType.Current
        };

        // Perform the comparison. Revisions will be added to the original document.
        original.Compare(edited, "AuthorInitials", DateTime.Now, options);

        // Save the result with revisions highlighted.
        original.Save("ComparisonResult.docx");
    }
}
