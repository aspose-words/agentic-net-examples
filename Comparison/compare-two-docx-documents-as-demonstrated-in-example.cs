using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Compare the documents, creating revisions in the original document.
        // "JD" will be used as the author initials for the revisions.
        docOriginal.Compare(docEdited, "JD", DateTime.Now);

        // Optional: iterate through the generated revisions and output their details.
        foreach (Revision rev in docOriginal.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Node type: {rev.ParentNode.NodeType}");
            Console.WriteLine($"\tChanged text: \"{rev.ParentNode.GetText()}\"");
        }

        // Accept all revisions to transform the original document into the edited version.
        docOriginal.Revisions.AcceptAll();

        // Save the resulting document.
        docOriginal.Save("ComparedResult.docx");
    }
}
