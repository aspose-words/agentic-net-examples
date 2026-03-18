using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the document that may contain tracked changes.
        Document doc = new Document("input.docx");

        // Accept every revision in the document.
        // This removes the revisions from the collection.
        doc.AcceptAllRevisions(); // Equivalent to doc.Revisions.AcceptAll();

        // Save the document now that all revision marks have been cleared.
        doc.Save("output.docx");
    }
}
