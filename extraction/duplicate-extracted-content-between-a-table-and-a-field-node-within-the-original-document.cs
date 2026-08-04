using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a table, some paragraphs, and a field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Intro paragraph before table.");

        // Insert a table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Paragraphs that will be between the table and the field.
        builder.Writeln("First paragraph between table and field.");
        builder.Writeln("Second paragraph between table and field.");

        // Insert a field and keep a reference to its containing paragraph.
        builder.InsertField("MERGEFIELD SampleField", "«SampleField»");
        Paragraph fieldParagraph = builder.CurrentParagraph;

        builder.Writeln("Paragraph after field.");

        // Save the initial document (optional, for inspection).
        const string inputPath = "sample.docx";
        doc.Save(inputPath);

        // Locate the first table in the document.
        Table table = doc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (table == null)
            throw new InvalidOperationException("Table not found in the document.");

        // Get the body that contains the nodes.
        Body body = doc.FirstSection.Body;

        // Collect nodes that lie between the table and the field paragraph (exclusive).
        List<Node> nodesToDuplicate = new List<Node>();
        Node current = table.NextSibling;
        while (current != null && current != fieldParagraph)
        {
            nodesToDuplicate.Add(current);
            current = current.NextSibling;
        }

        if (nodesToDuplicate.Count == 0)
            throw new InvalidOperationException("No content found between the table and the field.");

        // Clone the extracted nodes to preserve formatting.
        List<Node> clonedNodes = new List<Node>();
        foreach (Node node in nodesToDuplicate)
        {
            clonedNodes.Add(node.Clone(true));
        }

        // Insert the cloned nodes after the field's paragraph.
        Node referenceNode = fieldParagraph;
        foreach (Node clone in clonedNodes)
        {
            body.InsertAfter(clone, referenceNode);
            referenceNode = clone; // Update reference for next insertion.
        }

        // Save the resulting document.
        const string outputPath = "duplicated.docx";
        doc.Save(outputPath);

        // Validation: ensure the cloned nodes were inserted after the field paragraph.
        Node firstInserted = fieldParagraph.NextSibling;
        if (firstInserted == null || firstInserted != clonedNodes[0])
            throw new InvalidOperationException("Duplication validation failed: first cloned node not found after field.");

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        Console.WriteLine("Content between table and field duplicated successfully.");
    }
}
