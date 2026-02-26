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

        // Comparison can only be performed when both documents have no existing revisions.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Optional: configure comparison options.
            CompareOptions options = new CompareOptions
            {
                // Track changes at the word level.
                Granularity = Granularity.WordLevel,
                // Ignore case differences.
                IgnoreCaseChanges = true,
                // Use the edited document as the base for comparison.
                Target = ComparisonTargetType.New
            };

            // Perform the comparison. Revisions are added to docOriginal.
            docOriginal.Compare(docEdited, "JD", DateTime.Now, options);
        }

        // Save the result which now contains the revision markup.
        docOriginal.Save("ComparedResult.docx");
    }
}
