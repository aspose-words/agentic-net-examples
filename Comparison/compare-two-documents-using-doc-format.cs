using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the original and edited DOC files.
        string originalPath = "Original.doc";
        string editedPath = "Edited.doc";

        // Load the documents using the Document constructor (load rule).
        Document docOriginal = new Document(originalPath);
        Document docEdited = new Document(editedPath);

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The author name and current time are used for the revisions.
            docOriginal.Compare(docEdited, "Author", DateTime.Now);
        }

        // Save the comparison result to a new DOC file (save rule).
        string resultPath = "ComparisonResult.doc";
        docOriginal.Save(resultPath);
    }
}
