using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the source DOCM files and the result file.
        string originalPath = @"C:\Docs\Original.docm";
        string editedPath   = @"C:\Docs\Edited.docm";
        string resultPath   = @"C:\Docs\Compared.docm";

        // Load the two documents. The Document constructor handles the DOCM format automatically.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure that neither document contains existing revisions; otherwise Compare will throw.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform the comparison. The author initials and the current time are recorded in each revision.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // Save the original document, which now contains the revision markup, as a DOCM file.
        docOriginal.Save(resultPath);
    }
}
