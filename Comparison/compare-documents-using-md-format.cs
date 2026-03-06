using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparer
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("original.docx");
        Document docEdited = new Document("edited.docx");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. Revisions will be added to the original document.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Build a Markdown report of the revisions.
        StringBuilder markdown = new StringBuilder();
        markdown.AppendLine("# Document Comparison Report");
        markdown.AppendLine();

        foreach (Revision revision in docOriginal.Revisions)
        {
            // Each revision is listed with its type, node type, and the changed text.
            string revisionText = revision.ParentNode.GetText().Trim().Replace("\r", " ").Replace("\n", " ");
            markdown.AppendLine($"- **{revision.RevisionType}** in **{revision.ParentNode.NodeType}**: `{revisionText}`");
        }

        // Save the Markdown report to a file.
        File.WriteAllText("comparison.md", markdown.ToString());
    }
}
