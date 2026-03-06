using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original DOCM document.
        Document docOriginal = new Document("Original.docm");

        // Load the edited DOCM document to compare against.
        Document docEdited = new Document("Edited.docm");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "AuthorInitials", DateTime.Now);
        }

        // Optional: Accept all revisions so that docOriginal becomes identical to docEdited.
        docOriginal.Revisions.AcceptAll();

        // Save the resulting document (with revisions applied) to a new DOCM file.
        docOriginal.Save("Result.docm");
    }
}
