using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the original and edited documents.
        string originalPath = @"C:\Docs\Original.docx";
        string editedPath   = @"C:\Docs\Edited.docx";

        // Load the documents using the Document constructor (lifecycle rule).
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no revisions before comparison (required by the API).
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Configure comparison options (optional – here we compare everything).
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
            Target = ComparisonTargetType.New // Use the edited document as the target.
        };

        // Perform the comparison. The original document will receive Revision objects.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the result using the Document.Save method (lifecycle rule).
        string resultPath = @"C:\Docs\ComparisonResult.docx";
        docOriginal.Save(resultPath);
    }
}
