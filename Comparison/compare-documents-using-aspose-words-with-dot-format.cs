using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Set up comparison options (optional – here we compare everything).
        CompareOptions options = new CompareOptions
        {
            // Example: ignore case changes only.
            IgnoreCaseChanges = false,
            // Use the original document as the base for comparison.
            Target = ComparisonTargetType.Current
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, options);

        // Save the result as a Word template (.dot) – this demonstrates DOT format output.
        docOriginal.Save("ComparisonResult.dot", SaveFormat.Dot);
    }
}
