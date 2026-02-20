using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and the edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Configure comparison options (customize as needed).
        CompareOptions compareOptions = new CompareOptions
        {
            // Example: track changes at the word level and ignore case changes.
            Granularity = Granularity.WordLevel,
            IgnoreCaseChanges = true,
            // Use the edited document as the target for comparison.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the resulting document with revisions in DOCX format.
        docOriginal.Save("ComparedResult.docx");
    }
}
