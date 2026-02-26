using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace DocumentComparisonDemo
{
    class Program
    {
        static void Main()
        {
            // Load the original and the edited documents.
            Document docOriginal = new Document("Original.docx");
            Document docEdited   = new Document("Edited.docx");

            // Ensure both documents have no revisions before comparison.
            if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
                throw new InvalidOperationException("Documents must not contain revisions before comparison.");

            // Configure comparison options (customize as needed).
            CompareOptions compareOptions = new CompareOptions
            {
                // Example: ignore formatting changes and comments.
                IgnoreFormatting = true,
                IgnoreComments   = true,
                // Use the edited document as the target for comparison.
                Target = ComparisonTargetType.New
            };

            // Perform the comparison. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

            // Save the result document which now contains revision marks.
            docOriginal.Save("ComparedResult.docx");
        }
    }
}
