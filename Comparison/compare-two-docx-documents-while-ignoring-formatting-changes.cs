using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocs
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // The Compare method requires both documents to be revision‑free.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must not contain revisions before comparison.");

        // Configure comparison to ignore all formatting changes.
        CompareOptions options = new CompareOptions
        {
            IgnoreFormatting = true   // Formatting differences will not generate revisions.
            // All other options retain their default values (false).
        };

        // Perform the comparison. Revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, options);

        // Save the result which now contains revision marks.
        docOriginal.Save("ComparisonResult.docx");
    }
}
