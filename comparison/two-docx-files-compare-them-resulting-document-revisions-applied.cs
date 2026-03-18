using System;
using Aspose.Words;

namespace DocumentComparisonExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the source documents.
            const string originalPath = "Original.docx";
            const string editedPath   = "Edited.docx";

            // Load the two documents from disk.
            Document docOriginal = new Document(originalPath);
            Document docEdited   = new Document(editedPath);

            // Ensure both documents have no existing revisions before comparison.
            if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
            {
                // Compare the documents. All differences will be recorded as revisions
                // in the original document (docOriginal).
                docOriginal.Compare(docEdited, "Author", DateTime.Now);
            }

            // Save the document that now contains the revision markup.
            const string resultPath = "ComparedResult.docx";
            docOriginal.Save(resultPath);
        }
    }
}
