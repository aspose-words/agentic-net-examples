using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and the edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison options (customize as needed).
            CompareOptions compareOptions = new CompareOptions
            {
                // Example settings – adjust according to your requirements.
                IgnoreFormatting = false,
                IgnoreCaseChanges = false,
                Target = ComparisonTargetType.New
            };

            // Compare the documents. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);
        }

        // Save the document that now contains the revision marks.
        docOriginal.Save("ComparedResult.docx");

        // Output the number of revisions detected.
        Console.WriteLine($"Revisions count: {docOriginal.Revisions.Count}");
    }
}
