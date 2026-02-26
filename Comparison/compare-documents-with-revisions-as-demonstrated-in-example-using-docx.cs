using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocumentsWithRevisions
{
    static void Main()
    {
        // Path to the original and edited DOCX files.
        string originalPath = @"Original.docx";
        string editedPath   = @"Edited.docx";
        string resultPath   = @"ComparedResult.docx";

        // Load the two documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Optional: configure comparison options (e.g., track changes in headers/footers).
            CompareOptions options = new CompareOptions
            {
                // Set any flags you need; here we keep defaults.
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

            // Perform the comparison. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "JD", DateTime.Now, options);
        }

        // Save the document that now contains revision marks.
        docOriginal.Save(resultPath);
    }
}
