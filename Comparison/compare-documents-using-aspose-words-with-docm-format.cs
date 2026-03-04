using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the original and edited DOCM files.
        string originalPath = "Original.docm";
        string editedPath   = "Edited.docm";
        string resultPath   = "Compared.docm";

        // Load the documents from the file system.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Save the comparison result (original document now contains revisions) as a DOCM file.
        docOriginal.Save(resultPath);
    }
}
