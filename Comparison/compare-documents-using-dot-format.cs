using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The author name and timestamp are required.
            docOriginal.Compare(docEdited, "A", DateTime.Now);
        }

        // Build a DOT graph that represents each revision as a node.
        StringBuilder dot = new StringBuilder();
        dot.AppendLine("digraph Revisions {");
        dot.AppendLine("  node [shape=box];");

        for (int i = 0; i < docOriginal.Revisions.Count; i++)
        {
            var rev = docOriginal.Revisions[i];
            string nodeId = $"rev{i}";
            // Escape quotes and newlines for DOT label.
            string revText = rev.ParentNode.GetText()
                .Replace("\r", "")
                .Replace("\n", " ")
                .Replace("\"", "\\\"");
            string label = $"{rev.RevisionType}\\n{revText}";
            dot.AppendLine($"  {nodeId} [label=\"{label}\"];");

            // Connect nodes sequentially to visualize the order of revisions.
            if (i > 0)
                dot.AppendLine($"  rev{i - 1} -> {nodeId};");
        }

        dot.AppendLine("}");

        // Save the DOT representation to a file.
        File.WriteAllText("Revisions.dot", dot.ToString());

        // Accept all revisions and save the merged document (optional).
        docOriginal.Revisions.AcceptAll();
        docOriginal.Save("Merged.docx");
    }
}
