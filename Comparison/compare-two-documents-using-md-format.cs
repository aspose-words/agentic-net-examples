using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the original and edited documents.
        string originalPath = "Original.docx";
        string editedPath = "Edited.docx";
        string markdownOutputPath = "Comparison.md";

        // Load the documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited = new Document(editedPath);

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Perform the comparison. The original document will receive the revisions.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Build Markdown representation of the revisions.
        StringBuilder mdBuilder = new StringBuilder();
        mdBuilder.AppendLine("# Document Comparison Report");
        mdBuilder.AppendLine();
        mdBuilder.AppendLine($"*Compared on:* {DateTime.Now}");
        mdBuilder.AppendLine();
        mdBuilder.AppendLine("## Revisions");
        mdBuilder.AppendLine();

        foreach (Revision rev in docOriginal.Revisions)
        {
            // Revision type (Insert, Delete, Format, Move, etc.).
            string revType = rev.RevisionType.ToString();

            // Node type where the revision occurred.
            string nodeType = rev.ParentNode.NodeType.ToString();

            // Text of the affected node (trimmed to avoid excessive whitespace).
            string text = rev.ParentNode.GetText().Trim();

            // Escape pipe characters to keep table formatting safe.
            text = text.Replace("|", "\\|");

            mdBuilder.AppendLine($"- **{revType}** in **{nodeType}**: `{text}`");
        }

        // Write the Markdown to a file.
        File.WriteAllText(markdownOutputPath, mdBuilder.ToString());

        // Optionally, accept all revisions to transform the original into the edited version.
        // docOriginal.Revisions.AcceptAll();
        // docOriginal.Save("MergedResult.docx");
    }
}
