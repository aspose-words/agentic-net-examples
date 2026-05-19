using System;
using System.IO;
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

        // Paragraph before the table.
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

        // Paragraph between the table and the field with distinct formatting.
        builder.Font.Bold = true;
        builder.Writeln("Formatted paragraph between table and field.");
        builder.Font.Bold = false;

        // Insert a MERGEFIELD.
        builder.InsertField("MERGEFIELD SampleField", "SampleField");

        // Paragraph after the field.
        builder.Writeln("Paragraph after the field.");

        // Save the source document.
        const string sourcePath = "source.docx";
        doc.Save(sourcePath);

        // Load the document for processing.
        Document loaded = new Document(sourcePath);

        // Locate the first table in the document.
        Table targetTable = loaded.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (targetTable == null)
            throw new InvalidOperationException("Table not found in the document.");

        // Locate the first field in the document.
        if (loaded.Range.Fields.Count == 0)
            throw new InvalidOperationException("Field not found in the document.");
        Field targetField = loaded.Range.Fields[0];

        // Identify the paragraph that lies between the table and the field.
        Paragraph betweenParagraph = targetTable.NextSibling as Paragraph;
        if (betweenParagraph == null)
            throw new InvalidOperationException("Paragraph between table and field not found.");

        // Clone the paragraph to preserve original formatting.
        Paragraph clonedParagraph = (Paragraph)betweenParagraph.Clone(true);

        // Determine the paragraph that contains the field end node.
        Paragraph fieldParagraph = targetField.End.ParentNode as Paragraph;
        if (fieldParagraph == null)
            throw new InvalidOperationException("Field's containing paragraph not found.");

        // Insert the cloned paragraph after the field's paragraph.
        CompositeNode parent = fieldParagraph.ParentNode;
        parent.InsertAfter(clonedParagraph, fieldParagraph);

        // Validate that formatting (bold) was retained in the cloned paragraph.
        if (clonedParagraph.Runs.Count == 0 || !clonedParagraph.Runs[0].Font.Bold)
            throw new InvalidOperationException("Cloned paragraph formatting does not match the source.");

        // Save the resulting document.
        const string resultPath = "result.docx";
        loaded.Save(resultPath);

        // Verify that the result file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");
    }
}
