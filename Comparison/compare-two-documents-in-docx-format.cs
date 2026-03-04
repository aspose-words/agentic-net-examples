using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the two DOCX files to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // The Compare method requires that both documents have no existing revisions.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform the comparison.
            // "JD" is the author initials that will appear on the revision marks.
            // DateTime.Now supplies the timestamp for the revisions.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // Save the original document, which now contains the revision marks showing the differences.
        docOriginal.Save("ComparisonResult.docx");
    }
}
