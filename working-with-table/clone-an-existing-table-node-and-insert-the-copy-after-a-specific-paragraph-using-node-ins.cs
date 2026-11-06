using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph before the table.
        builder.Writeln("Paragraph before the table.");

        // Build a simple 2‑row table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Add a paragraph after the original table.
        builder.Writeln("Paragraph after the original table.");

        // Save the initial document (optional, just to have a file).
        string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Locate the paragraph after which we want to insert the cloned table.
        // Here we use the first paragraph ("Paragraph before the table.").
        Paragraph targetParagraph = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[0];

        // Find the first table in the document.
        Table originalTable = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Clone the table (deep clone).
        Table clonedTable = (Table)originalTable.Clone(true);

        // Insert the cloned table after the target paragraph.
        // InsertAfter is called on the parent node of the reference paragraph.
        targetParagraph.ParentNode.InsertAfter(clonedTable, targetParagraph);

        // Save the resulting document.
        string resultPath = "ClonedTable.docx";
        doc.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
        {
            throw new Exception("The output document was not created.");
        }
    }
}
