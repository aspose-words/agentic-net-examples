using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the two DOCX files to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Perform the comparison. The original document will receive revisions for any differences.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // If the revisions count is zero, the documents are identical.
        bool areIdentical = docOriginal.Revisions.Count == 0;

        Console.WriteLine($"Documents are {(areIdentical ? "identical" : "different")}.");

        // Optionally, save the result document (it will be identical to the edited one if no revisions were added).
        docOriginal.Save("ComparisonResult.docx");
    }
}
