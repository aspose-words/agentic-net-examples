using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original DOT (template) document.
        Document docOriginal = new Document("OriginalTemplate.dot");

        // Load the edited DOT document that we want to compare against.
        Document docEdited = new Document("EditedTemplate.dot");

        // Ensure both documents have no pending revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
        {
            docOriginal.Revisions.RejectAll();
            docEdited.Revisions.RejectAll();
        }

        // Configure comparison options.
        // Example: ignore formatting changes and track differences at the word level.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true,
            Granularity = Granularity.WordLevel,
            Target = ComparisonTargetType.New // Use the edited document as the target.
        };

        // Perform the comparison.
        // After this call, docOriginal will contain Revision objects for each difference.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Accept all revisions so that docOriginal becomes identical to docEdited.
        docOriginal.Revisions.AcceptAll();

        // Save the resulting document that reflects the comparison outcome.
        docOriginal.Save("ComparisonResult.docx");
    }
}
