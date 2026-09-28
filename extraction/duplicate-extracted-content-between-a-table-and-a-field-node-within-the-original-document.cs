using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a table, some paragraphs, and a PAGE field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph before table.");

        // Insert a table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("Paragraph between table and field.");
        builder.Writeln("Another paragraph between table and field.");

        // Insert a PAGE field.
        builder.InsertField("PAGE", null);
        builder.Writeln("Paragraph after field.");

        // Save the source document.
        string sourcePath = "source.docx";
        doc.Save(sourcePath);

        // Load the document for processing.
        Document loaded = new Document(sourcePath);

        // Locate the first table.
        Table table = loaded.GetChildNodes(NodeType.Table, true)
            .Cast<Table>()
            .FirstOrDefault();

        if (table == null)
            throw new InvalidOperationException("Table not found in the document.");

        // Locate the first PAGE field start node.
        FieldStart fieldStart = loaded.GetChildNodes(NodeType.FieldStart, true)
            .Cast<FieldStart>()
            .FirstOrDefault(fs => fs.FieldType == FieldType.FieldPage);

        if (fieldStart == null)
            throw new InvalidOperationException("PAGE field not found in the document.");

        // Determine the range of nodes between the table and the field (exclusive).
        List<Node> nodesBetween = new List<Node>();
        Node current = table.NextSibling;
        while (current != null && current != fieldStart)
        {
            nodesBetween.Add(current);
            current = current.NextSibling;
        }

        if (nodesBetween.Count == 0)
            throw new InvalidOperationException("No content found between the table and the field.");

        // Clone the extracted nodes and insert them after the paragraph that contains the field.
        Paragraph fieldParagraph = fieldStart.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (fieldParagraph == null)
            throw new InvalidOperationException("Field is not inside a paragraph.");

        CompositeNode parent = fieldParagraph.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Unable to locate a valid parent for insertion.");

        Node referenceNode = fieldParagraph;

        foreach (Node node in nodesBetween)
        {
            Node cloned = node.Clone(true);
            parent.InsertAfter(cloned, referenceNode);
            referenceNode = cloned; // Subsequent clones are inserted after the previous one.
        }

        // Save the modified document.
        string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Validation: ensure the duplicated content exists after the field.
        Document resultDoc = new Document(outputPath);
        FieldStart resultFieldStart = resultDoc.GetChildNodes(NodeType.FieldStart, true)
            .Cast<FieldStart>()
            .FirstOrDefault(fs => fs.FieldType == FieldType.FieldPage);

        if (resultFieldStart == null)
            throw new InvalidOperationException("PAGE field not found in the result document.");

        Paragraph resultFieldParagraph = resultFieldStart.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (resultFieldParagraph == null)
            throw new InvalidOperationException("Field paragraph not found in the result document.");

        // Collect texts of the original nodes (trimmed for comparison).
        List<string> originalTexts = nodesBetween
            .Select(n => n.GetText().Trim())
            .Where(t => !string.IsNullOrEmpty(t))
            .ToList();

        // Walk nodes after the field paragraph and count matches.
        int duplicatedCount = 0;
        Node nodeAfter = resultFieldParagraph.NextSibling;
        while (nodeAfter != null && duplicatedCount < originalTexts.Count)
        {
            string txt = nodeAfter.GetText().Trim();
            if (originalTexts.Contains(txt))
                duplicatedCount++;

            nodeAfter = nodeAfter.NextSibling;
        }

        if (duplicatedCount < originalTexts.Count)
            throw new InvalidOperationException("Duplicated content was not inserted correctly.");

        Console.WriteLine("Extraction and duplication completed successfully.");
    }
}
