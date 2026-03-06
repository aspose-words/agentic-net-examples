using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocumentsIgnoringFormatting
{
    static void Main()
    {
        // Paths to the original and edited documents.
        string originalPath = @"C:\Docs\Original.docx";
        string editedPath   = @"C:\Docs\Edited.docx";
        string resultPath   = @"C:\Docs\ComparisonResult.docx";

        // Load the documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no existing revisions (required by Aspose.Words).
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Configure comparison options to ignore formatting changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true,               // <-- ignore all formatting differences
            CompareMoves = false,                  // optional: do not track moved text
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New      // compare against the edited document
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the document that now contains the revision markup.
        docOriginal.Save(resultPath);
    }
}
