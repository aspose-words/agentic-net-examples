using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocumentsToDot
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. All differences will be stored as revisions in docOriginal.
            docOriginal.Compare(docEdited, "Author", DateTime.Now);
        }

        // Save the comparison result in GraphViz DOT format.
        // The DOT file visualizes the revision hierarchy.
        docOriginal.Save("ComparisonResult.dot", SaveFormat.Dot);
    }
}
