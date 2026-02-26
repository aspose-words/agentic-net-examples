using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Both documents must be revision‑free before invoking Compare.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison options (optional).
            CompareOptions options = new CompareOptions
            {
                // Track changes at the word level.
                Granularity = Granularity.WordLevel,
                // Show changes relative to the edited document.
                Target = ComparisonTargetType.New
            };

            // Compare the documents. Revisions are added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now, options);
        }

        // Save the result – the document now contains revision marks.
        docOriginal.Save("ComparisonResult.docx");
    }
}
