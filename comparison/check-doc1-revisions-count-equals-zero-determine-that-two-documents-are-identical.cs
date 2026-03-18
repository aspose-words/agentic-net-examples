using System;
using Aspose.Words;

class CompareDocuments
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Verify that neither document contains revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Compare the documents. Any differences will be recorded as revisions in docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // If the revision count is zero after comparison, the documents are identical.
        bool areIdentical = docOriginal.Revisions.Count == 0;

        Console.WriteLine(areIdentical ? "Documents are identical." : "Documents differ.");

        // Save the comparison result (contains revisions if differences were found).
        docOriginal.Save("ComparisonResult.docx");
    }
}
