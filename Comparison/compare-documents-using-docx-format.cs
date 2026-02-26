using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocs
{
    static void Main()
    {
        // Load the original document.
        Document docOriginal = new Document("Original.docx");

        // Load the edited document to compare against.
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no existing revisions; otherwise Compare will throw.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison options (optional – all false means track all changes).
            CompareOptions compareOptions = new CompareOptions
            {
                CompareMoves = false,
                IgnoreFormatting = false,
                IgnoreCaseChanges = false,
                IgnoreComments = false,
                IgnoreTables = false,
                IgnoreFields = false,
                IgnoreFootnotes = false,
                IgnoreTextboxes = false,
                IgnoreHeadersAndFooters = false,
                Target = ComparisonTargetType.New
            };

            // Perform the comparison. Revisions are added to docOriginal.
            docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, compareOptions);
        }

        // Save the document that now contains the tracked changes (revisions).
        docOriginal.Save("ComparisonResult.docx");
    }
}
