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

        // The Compare method requires that both documents have no existing revisions.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison to ignore formatting changes.
            CompareOptions compareOptions = new CompareOptions
            {
                IgnoreFormatting = true,                 // <-- ignore all formatting differences
                Target = ComparisonTargetType.New       // compare against the edited document
                // Other options remain at their default values (false).
            };

            // Perform the comparison. Revisions describing the differences are added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);
        }

        // Save the result document which now contains the revision marks.
        docOriginal.Save("ComparisonResult.docx");
    }
}
