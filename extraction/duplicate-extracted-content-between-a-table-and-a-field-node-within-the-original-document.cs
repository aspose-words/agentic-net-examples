using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // -------------------- Create sample document --------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph before the table.");

        // Insert a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("A1");
        builder.InsertCell();
        builder.Write("B1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("A2");
        builder.InsertCell();
        builder.Write("B2");
        builder.EndRow();
        builder.EndTable();

        // Paragraph that will be extracted and duplicated.
        builder.Writeln("Between content that will be duplicated.");

        // Insert a MERGEFIELD as the field node.
        builder.InsertField("MERGEFIELD SampleField", "FieldResult");

        builder.Writeln("Paragraph after the field.");

        // Save the source document (optional, just for inspection).
        const string sourcePath = "source.docx";
        doc.Save(sourcePath);

        // -------------------- Load document and locate nodes --------------------
        Document loadedDoc = new Document(sourcePath);

        // Locate the first table in the document.
        Table targetTable = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (targetTable == null)
            throw new InvalidOperationException("Table not found in the document.");

        // Locate the first field in the document.
        Field field = loadedDoc.Range.Fields[0];
        if (field == null)
            throw new InvalidOperationException("Field not found in the document.");

        // Get the paragraph that contains the field.
        Paragraph fieldParagraph = field.Start.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (fieldParagraph == null)
            throw new InvalidOperationException("Field paragraph not found.");

        // Get the body that holds the block‑level nodes.
        Body body = loadedDoc.FirstSection.Body;

        // -------------------- Collect nodes between the table and the field paragraph --------------------
        List<Node> nodesToDuplicate = new List<Node>();
        Node curNode = targetTable.NextSibling;
        while (curNode != null && curNode != fieldParagraph)
        {
            nodesToDuplicate.Add(curNode);
            curNode = curNode.NextSibling;
        }

        if (nodesToDuplicate.Count == 0)
            throw new InvalidOperationException("No content found between the table and the field.");

        // -------------------- Duplicate the collected nodes after the field paragraph --------------------
        Node insertionPoint = fieldParagraph;
        foreach (Node node in nodesToDuplicate)
        {
            // Clone the node (deep clone) and insert it after the current insertion point.
            Node clonedNode = node.Clone(true);
            body.InsertAfter(clonedNode, insertionPoint);
            insertionPoint = clonedNode;
        }

        // Save the resulting document.
        const string resultPath = "result.docx";
        loadedDoc.Save(resultPath);

        // Validation: ensure the duplicated paragraph exists twice.
        int duplicatedParagraphCount = 0;
        foreach (Paragraph para in loadedDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.GetText().Trim() == "Between content that will be duplicated.")
                duplicatedParagraphCount++;
        }

        if (duplicatedParagraphCount != 2)
            throw new InvalidOperationException("The content was not duplicated correctly.");

        Console.WriteLine("Content duplicated successfully. Result saved to " + resultPath);
    }
}
