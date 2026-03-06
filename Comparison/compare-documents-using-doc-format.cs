using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and the edited DOC documents.
        Document docOriginal = new Document("Original.doc");
        Document docEdited = new Document("Edited.doc");

        // The Compare method requires both documents to be revision‑free.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The original document will receive Revision objects
            // describing the differences. "JD" is the author initials, and the current
            // date/time is used for the revision metadata.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // If you want the original document to become identical to the edited one,
        // accept all revisions that were created during the comparison.
        docOriginal.Revisions.AcceptAll();

        // Save the result as a DOC file. The file will contain the tracked changes
        // (if any) or the fully accepted content.
        docOriginal.Save("ComparisonResult.doc");
    }
}
