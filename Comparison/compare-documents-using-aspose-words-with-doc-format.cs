using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and the edited DOC files.
        Document docOriginal = new Document("Original.doc");
        Document docEdited   = new Document("Edited.doc");

        // Ensure that neither document already contains revisions.
        // Aspose.Words throws an exception if revisions exist before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents.
            // The revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Accept all revisions so that docOriginal becomes identical to docEdited.
        docOriginal.Revisions.AcceptAll();

        // Save the result as a DOC file.
        docOriginal.Save("ComparedResult.doc");
    }
}
