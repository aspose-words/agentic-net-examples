using System;
using Aspose.Words;

class CompareDocumentSelf
{
    static void Main()
    {
        // Path to the source document.
        string inputPath = "input.docx";

        // Load the document from disk.
        Document doc = new Document(inputPath);

        // Ensure the document has no revisions before comparison.
        if (doc.HasRevisions)
            throw new InvalidOperationException("Document already contains revisions.");

        // Compare the document with itself.
        // This should not generate any revisions because the contents are identical.
        doc.Compare(doc, "SelfCompare", DateTime.Now);

        // Verify that no revisions were created.
        if (doc.Revisions.Count != 0)
            throw new InvalidOperationException($"Unexpected revisions count: {doc.Revisions.Count}");

        // Save the (unchanged) document back to disk if needed.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
