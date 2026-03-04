using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and edited documents from MHTML files.
        Document docOriginal = new Document("Original.mht");
        Document docEdited = new Document("Edited.mht");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. "JD" will be used as the author initials for the revisions.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // The original document now contains revision marks that represent the differences.
        // Optionally, accept all revisions to make the original identical to the edited version:
        // docOriginal.Revisions.AcceptAll();

        // Save the comparison result (including revision marks) back to MHTML format.
        docOriginal.Save("ComparisonResult.mht");
    }
}
