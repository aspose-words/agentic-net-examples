using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and edited documents from disk.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Set up comparison options (optional). Here we track changes at the word level
        // and ignore formatting differences.
        CompareOptions compareOptions = new CompareOptions
        {
            Granularity = Granularity.WordLevel,
            IgnoreFormatting = true,
            Target = ComparisonTargetType.Current   // Use the original document as the base.
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, compareOptions);

        // Save the document that now contains revision marks.
        docOriginal.Save("ComparisonResult.docx");
    }
}
