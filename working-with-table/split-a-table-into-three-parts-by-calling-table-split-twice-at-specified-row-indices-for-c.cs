using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with 9 rows, each row contains a single cell with text.
        builder.StartTable();
        for (int i = 1; i <= 9; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i}");
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the first (and only) table in the document.
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
            throw new InvalidOperationException("Table was not created.");

        // Store the rows for easier processing.
        var rows = originalTable.Rows.Cast<Row>().ToList();

        // Create three new tables that are clones of the original table structure (without rows).
        Table firstPart = (Table)originalTable.Clone(false);
        Table middlePart = (Table)originalTable.Clone(false);
        Table lastPart = (Table)originalTable.Clone(false);

        // Add rows 0‑2 to the first part.
        for (int i = 0; i <= 2; i++)
            firstPart.Rows.Add((Row)rows[i].Clone(true));

        // Add rows 3‑5 to the middle part.
        for (int i = 3; i <= 5; i++)
            middlePart.Rows.Add((Row)rows[i].Clone(true));

        // Add rows 6‑8 to the last part.
        for (int i = 6; i <= 8; i++)
            lastPart.Rows.Add((Row)rows[i].Clone(true));

        // Insert the new tables into the document at the position of the original table.
        CompositeNode parent = originalTable.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Unable to locate a valid parent node for insertion.");

        // Insert after the original table in reverse order so the final order is correct.
        parent.InsertAfter(lastPart, originalTable);
        parent.InsertAfter(middlePart, originalTable);
        parent.InsertAfter(firstPart, originalTable);

        // Remove the original table.
        originalTable.Remove();

        // Validate that the document now contains three separate tables.
        NodeCollection allTables = doc.GetChildNodes(NodeType.Table, true);
        if (allTables.Count != 3)
            throw new InvalidOperationException($"Expected 3 tables after splitting, but found {allTables.Count}.");

        // Output row counts of each resulting table to the console.
        Console.WriteLine($"First part rows: {((Table)allTables[0]).Rows.Count}");
        Console.WriteLine($"Second part rows: {((Table)allTables[1]).Rows.Count}");
        Console.WriteLine($"Third part rows: {((Table)allTables[2]).Rows.Count}");

        // Save the resulting document.
        string outputPath = "SplitTable.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
