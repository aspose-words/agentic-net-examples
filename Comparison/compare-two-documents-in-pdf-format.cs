using System;
using Aspose.Words;

class ComparePdfDocuments
{
    static void Main()
    {
        // Load the first PDF document (the original version).
        Document originalDoc = new Document("Original.pdf");

        // Load the second PDF document (the version to compare against).
        Document editedDoc = new Document("Edited.pdf");

        // Both documents must not contain revisions before comparison.
        if (originalDoc.Revisions.Count == 0 && editedDoc.Revisions.Count == 0)
        {
            // Compare the documents. Revisions describing the differences
            // will be added to the original document.
            originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);
        }

        // Save the comparison result (including revisions) as a new PDF file.
        originalDoc.Save("ComparisonResult.pdf");
    }
}
