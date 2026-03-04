using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and the edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Configure comparison options to ignore any changes in headers and footers.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreHeadersAndFooters = true,
            // Keep other options at their defaults.
            Target = ComparisonTargetType.New
        };

        // Compare the documents. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, compareOptions);

        // Save the comparison result.
        docOriginal.Save("ComparisonResult.docx");
    }
}
