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
        // ------------------------------------------------------------
        // Create a sample document containing a table, a paragraph,
        // and a MERGEFIELD. This document will be used as the source.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Intro paragraph before table.");

        // Insert a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Paragraph that will be between the table and the field.
        builder.Writeln("Paragraph that lies between the table and the field.");

        // Insert a MERGEFIELD.
        builder.InsertField("MERGEFIELD MyField \\* MERGEFORMAT", "MyField");

        builder.Writeln("Paragraph after field.");

        // Save the source document to a deterministic local file.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // ------------------------------------------------------------
        // Load the document for processing.
        // ------------------------------------------------------------
        Document doc = new Document(sourcePath);

        // Locate the first table in the document.
        Table table = doc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (table == null)
            throw new InvalidOperationException("Table not found in the document.");

        // Locate the first field in the document.
        if (doc.Range.Fields.Count == 0)
            throw new InvalidOperationException("No fields found in the document.");
        Field field = doc.Range.Fields[0];
        Node fieldStartNode = field.Start; // Correct property for the start node.
        if (fieldStartNode == null)
            throw new InvalidOperationException("Field start node is missing.");

        // Determine the nodes that lie between the table and the field.
        List<Node> nodesBetween = new List<Node>();
        Node current = table.NextSibling;
        while (current != null && current != fieldStartNode)
        {
            nodesBetween.Add(current);
            current = current.NextSibling;
        }

        if (nodesBetween.Count == 0)
            throw new InvalidOperationException("No content found between the table and the field.");

        // Duplicate the extracted nodes after the field's containing paragraph.
        Paragraph fieldParagraph = fieldStartNode.ParentNode as Paragraph;
        if (fieldParagraph == null)
            throw new InvalidOperationException("Field is not inside a paragraph.");

        Node insertionPoint = fieldParagraph;
        foreach (Node node in nodesBetween)
        {
            Node clonedNode = node.Clone(true);
            // Insert after the current insertion point.
            insertionPoint.ParentNode.InsertAfter(clonedNode, insertionPoint);
            insertionPoint = clonedNode;
        }

        // Save the resulting document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // ------------------------------------------------------------
        // Validation: ensure the duplicated content exists at least twice.
        // ------------------------------------------------------------
        Document resultDoc = new Document(outputPath);
        int duplicatedCount = 0;
        foreach (Paragraph para in resultDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.GetText().Contains("Paragraph that lies between the table and the field."))
                duplicatedCount++;
        }

        if (duplicatedCount < 2)
            throw new InvalidOperationException("Duplicated content was not correctly inserted.");

        // Optional clean‑up (commented out to keep files for inspection).
        // File.Delete(sourcePath);
        // File.Delete(outputPath);
    }
}
