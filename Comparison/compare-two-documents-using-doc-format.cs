using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and edited DOC files.
        Document docOriginal = new Document("Original.doc");
        Document docEdited   = new Document("Edited.doc");

        // Documents must not contain revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the two documents. Revisions are added to docOriginal.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // List all revisions that were created by the comparison.
        foreach (Revision rev in docOriginal.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Node type: {rev.ParentNode.NodeType}");
            Console.WriteLine($"\tChanged text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Accept all revisions so that docOriginal becomes identical to docEdited.
        docOriginal.Revisions.AcceptAll();

        // Save the resulting document in DOC format.
        docOriginal.Save("Merged.doc");
    }
}
