using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the original, edited and DOT output files.
        string originalPath = "Original.docx";
        string editedPath   = "Edited.docx";
        string dotPath      = "Comparison.dot";

        // Load the two documents that will be compared.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Documents must not contain revisions before a comparison is performed.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Perform the comparison. The revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Build a simple DOT graph that lists each revision as a node.
        using (StreamWriter writer = new StreamWriter(dotPath))
        {
            writer.WriteLine("digraph DocumentComparison {");
            writer.WriteLine("  node [shape=box];");

            for (int i = 0; i < docOriginal.Revisions.Count; i++)
            {
                Revision rev = docOriginal.Revisions[i];
                string nodeId = $"rev{i}";
                // Escape quotes and newlines for DOT label.
                string label = rev.RevisionType.ToString()
                                 .Replace("\"", "\\\"") + "\\n" +
                                 rev.ParentNode.GetText()
                                 .Replace("\r", "")
                                 .Replace("\n", " ")
                                 .Replace("\"", "\\\"");
                writer.WriteLine($"  {nodeId} [label=\"{label}\"];");
            }

            writer.WriteLine("}");
        }

        // Optionally save the revised document for further inspection.
        docOriginal.Save("Revised.docx");
    }
}
