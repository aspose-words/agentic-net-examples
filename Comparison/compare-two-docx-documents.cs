using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace DocumentComparisonDemo
{
    class Program
    {
        static void Main()
        {
            // Load the original and edited documents.
            Document docOriginal = new Document("Original.docx");
            Document docEdited = new Document("Edited.docx");

            // Ensure both documents have no existing revisions before comparison.
            if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
                throw new InvalidOperationException("Documents must not contain revisions before comparison.");

            // Optional: configure comparison options (e.g., ignore formatting, compare moves, etc.).
            CompareOptions compareOptions = new CompareOptions
            {
                CompareMoves = false,
                IgnoreFormatting = false,
                IgnoreCaseChanges = false,
                IgnoreComments = false,
                IgnoreTables = false,
                IgnoreFields = false,
                IgnoreFootnotes = false,
                IgnoreTextboxes = false,
                IgnoreHeadersAndFooters = false,
                Target = ComparisonTargetType.New // Use the edited document as the target.
            };

            // Perform the comparison. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

            // Save the document that now contains the revision markup.
            docOriginal.Save("ComparedResult.docx");
        }
    }
}
