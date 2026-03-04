using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document original = new Document("Original.docx");
        Document edited   = new Document("Edited.docx");

        // The Compare method requires that both documents have no existing revisions.
        if (original.Revisions.Count != 0 || edited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Perform the comparison. All differences will be recorded as revisions in the original document.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Enumerate the revisions and write their details to the console.
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Node type: {rev.ParentNode.NodeType}");
            Console.WriteLine($"\tChanged text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the resulting document (which now contains the tracked changes) to a new file.
        original.Save("ComparedResult.docx");
    }
}
