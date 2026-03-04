using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDotm
{
    static void Main()
    {
        // Load the original and edited DOTM documents.
        Document docOriginal = new Document("Original.dotm");
        Document docEdited = new Document("Edited.dotm");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // Optionally accept all revisions to make docOriginal identical to docEdited.
        // docOriginal.Revisions.AcceptAll();

        // Save the result of the comparison as a DOTM file.
        docOriginal.Save("ComparisonResult.dotm");
    }
}
