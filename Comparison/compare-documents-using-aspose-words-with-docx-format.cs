using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original document from a DOCX file.
        Document original = new Document("Original.docx");

        // Load the edited document from a DOCX file.
        Document edited = new Document("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (original.Revisions.Count != 0 || edited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Set up comparison options (optional – customize as needed).
        CompareOptions options = new CompareOptions
        {
            // Example: ignore formatting changes.
            IgnoreFormatting = true,
            // Example: track changes at the word level.
            Granularity = Granularity.WordLevel,
            // Use the edited document as the base for comparison.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to the original document.
        original.Compare(edited, "Comparer", DateTime.Now, options);

        // Save the result (original document now contains revision marks).
        original.Save("ComparisonResult.docx");
    }
}
