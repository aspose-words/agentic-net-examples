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

        // Ensure both documents have no existing revisions before performing the comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform the comparison. The original document will receive revision marks.
            // "Comparer" is the author name that will appear on the revisions.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Save the result (the original document now contains the revisions) in WORDML format.
        // WORDML is the XML representation of a Word document (SaveFormat.WordML).
        docOriginal.Save("ComparisonResult.xml", SaveFormat.WordML);
    }
}
