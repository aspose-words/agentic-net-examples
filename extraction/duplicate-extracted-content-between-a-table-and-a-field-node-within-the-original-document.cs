using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // Create a sample document with a table, paragraphs and a field.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph before the table.");

        // 1x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table cell 1");
        builder.EndRow();
        builder.EndTable();

        // Paragraphs that lie between the table and the field.
        builder.Writeln("First paragraph between table and field.");
        builder.Writeln("Second paragraph between table and field.");

        // Insert a PAGE field.
        builder.InsertField("PAGE");

        // Paragraph after the field.
        builder.Writeln("Paragraph after the field.");

        // -------------------------------------------------
        // Duplicate the content that exists between the table and the field.
        // -------------------------------------------------

        // Locate the first table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table not found.");

        // Locate the first field.
        Field field = doc.Range.Fields[0];
        if (field == null)
            throw new InvalidOperationException("Field not found.");

        // The body that contains the table, the paragraphs, and the field.
        CompositeNode body = table.ParentNode as CompositeNode;
        if (body == null)
            throw new InvalidOperationException("Unable to locate the body node.");

        // Collect nodes that are positioned after the table and before the field start.
        List<Node> nodesBetween = new List<Node>();
        bool insideRange = false;

        foreach (Node child in body.GetChildNodes(NodeType.Any, false))
        {
            if (child == table)
            {
                // Start collecting after the table node.
                insideRange = true;
                continue;
            }

            // The field start node marks the end of the range we want to copy.
            if (child == field.Start)
            {
                insideRange = false;
                break;
            }

            if (insideRange)
                nodesBetween.Add(child);
        }

        if (nodesBetween.Count == 0)
            throw new InvalidOperationException("No content found between the table and the field.");

        // Prepare an importer that preserves the original formatting.
        NodeImporter importer = new NodeImporter(doc, doc, ImportFormatMode.KeepSourceFormatting);

        // Insert the duplicated nodes after the paragraph that contains the field end.
        // The field end is an inline node inside a paragraph, so we insert after that paragraph.
        Paragraph fieldParagraph = field.End.ParentNode as Paragraph;
        if (fieldParagraph == null)
            throw new InvalidOperationException("Unable to locate the paragraph containing the field.");

        Node insertionPoint = fieldParagraph;

        foreach (Node sourceNode in nodesBetween)
        {
            // Import (deep clone) the node into the same document.
            Node clonedNode = importer.ImportNode(sourceNode, true);
            body.InsertAfter(clonedNode, insertionPoint);
            insertionPoint = clonedNode; // Update insertion point for the next node.
        }

        // Save the resulting document.
        const string outputPath = "DuplicatedContent.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
