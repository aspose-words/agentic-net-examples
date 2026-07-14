using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Intro paragraph.
        builder.Writeln("Intro paragraph.");

        // Insert a table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // Content that will be extracted (two paragraphs).
        builder.Writeln("Extracted paragraph 1.");
        builder.Writeln("Extracted paragraph 2.");

        // Insert a MERGEFIELD (field node).
        builder.InsertField("MERGEFIELD SampleField");

        // Closing paragraph.
        builder.Writeln("After field paragraph.");

        // Save the original document (optional, for inspection).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        doc.Save(originalPath);

        // Locate the first table in the document.
        Table targetTable = doc.GetChildNodes(NodeType.Table, true).OfType<Table>().FirstOrDefault();
        if (targetTable == null)
            throw new InvalidOperationException("Table not found in the document.");

        // Locate the first field start node (the MERGEFIELD we inserted).
        FieldStart fieldStart = doc.GetChildNodes(NodeType.FieldStart, true).OfType<FieldStart>().FirstOrDefault();
        if (fieldStart == null)
            throw new InvalidOperationException("Field start not found in the document.");

        // Get the corresponding field end node via the Field object.
        Field field = fieldStart.GetField();
        FieldEnd fieldEnd = field?.End;
        if (fieldEnd == null)
            throw new InvalidOperationException("Field end not found in the document.");

        // The body that contains the nodes.
        Body body = doc.FirstSection.Body;

        // Collect block-level nodes that appear after the table and before the paragraph that holds the field start.
        var nodesToDuplicate = new System.Collections.Generic.List<Node>();
        Node current = targetTable.NextSibling;
        while (current != null && !(current is Paragraph para && para.GetChildNodes(NodeType.FieldStart, true).Contains(fieldStart)))
        {
            nodesToDuplicate.Add(current);
            current = current.NextSibling;
        }

        if (!nodesToDuplicate.Any())
            throw new InvalidOperationException("No nodes found between the table and the field.");

        // Duplicate each extracted node and insert after the paragraph that contains the field.
        // fieldEnd is an inline node inside a paragraph; we need the paragraph as the insertion point.
        Paragraph fieldParagraph = fieldEnd.ParentNode as Paragraph;
        if (fieldParagraph == null)
            throw new InvalidOperationException("Field end does not have a parent paragraph.");

        Node insertionPoint = fieldParagraph;
        foreach (Node node in nodesToDuplicate)
        {
            // Clone the node deeply to preserve formatting.
            Node clonedNode = node.Clone(true);
            // Insert the cloned block node after the field's paragraph.
            body.InsertAfter(clonedNode, insertionPoint);
            insertionPoint = clonedNode; // Update insertion point so duplicates stay in order.
        }

        // Save the resulting document.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "result.docx");
        doc.Save(resultPath);

        // Validate that the result file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");

        Console.WriteLine("Document processing completed successfully.");
    }
}
