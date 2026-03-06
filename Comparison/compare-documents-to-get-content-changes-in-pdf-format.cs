using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison options.
            CompareOptions compareOptions = new CompareOptions
            {
                // Track changes at the word level (can be set to CharLevel if needed).
                Granularity = Granularity.WordLevel,
                // Use the edited document as the target so revisions are stored in the original.
                Target = ComparisonTargetType.New,
                // Example: do not ignore any element types.
                IgnoreFormatting = false,
                IgnoreCaseChanges = false,
                IgnoreComments = false,
                IgnoreTables = false,
                IgnoreFields = false,
                IgnoreFootnotes = false,
                IgnoreTextboxes = false,
                IgnoreHeadersAndFooters = false,
                CompareMoves = false
            };

            // Perform the comparison. Revisions describing the differences are added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);
        }

        // Save the document with revisions to PDF. The PDF will contain the tracked changes.
        docOriginal.Save("ComparisonResult.pdf", SaveFormat.Pdf);
    }
}
