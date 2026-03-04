using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform the comparison. The author name and timestamp are required.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Build a GraphViz DOT representation of the revisions created by the comparison.
        StringBuilder dot = new StringBuilder();
        dot.AppendLine("digraph Revisions {");
        dot.AppendLine("  node [shape=box];");

        int revIndex = 0;
        foreach (Revision rev in docOriginal.Revisions)
        {
            // Create a unique node identifier.
            string nodeId = $"rev{revIndex}";

            // Prepare a label containing revision type, node type and the changed text.
            // Escape double quotes to keep the DOT syntax valid.
            string changedText = rev.ParentNode.GetText()
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\"", "\\\"")
                .Trim();

            string label = $"{rev.RevisionType}\\n{rev.ParentNode.NodeType}\\n{changedText}";
            dot.AppendLine($"  {nodeId} [label=\"{label}\"];");
            revIndex++;
        }

        dot.AppendLine("}");

        // Save the DOT graph to a file.
        File.WriteAllText("Revisions.dot", dot.ToString());

        // Optionally, save the original document now containing the revisions.
        docOriginal.Save("Original_With_Revisions.docx");
    }
}
