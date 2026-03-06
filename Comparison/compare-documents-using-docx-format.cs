using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocs
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");   // create/load rule
        Document docEdited   = new Document("Edited.docx");    // create/load rule

        // Ensure both documents have no existing revisions; otherwise Compare will throw.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison options (optional – all false means track every change).
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

            // Compare the documents. Revisions are added to docOriginal.
            docOriginal.Compare(docEdited, "AuthorName", DateTime.Now, compareOptions);
        }

        // Save the result document containing the tracked changes.
        docOriginal.Save("ComparisonResult.docx"); // save rule
    }
}
