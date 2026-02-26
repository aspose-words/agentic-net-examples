using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Documents must not contain revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform the comparison. The original document will receive Revision objects
            // that highlight the differences. "AB" are the author initials for the revisions.
            docOriginal.Compare(docEdited, "AB", DateTime.Now);
        }

        // Save the document that now contains the tracked changes.
        docOriginal.Save("ComparisonResult.docx");
    }
}
