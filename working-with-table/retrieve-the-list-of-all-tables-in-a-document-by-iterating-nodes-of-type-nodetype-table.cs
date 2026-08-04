using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define a file name for the sample document.
        string filePath = "SampleDocument.docx";

        // -----------------------------------------------------------------
        // 1. Create a new blank document and add a simple 2x2 table.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph before the table.
        builder.Writeln("This is a sample document containing tables.");

        // Start the table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save(filePath);

        // Verify that the file was created.
        if (!File.Exists(filePath))
            throw new Exception($"Failed to create the document at '{filePath}'.");

        // -----------------------------------------------------------------
        // 2. Load the document (could reuse the same instance, but loading
        //    demonstrates the typical workflow).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(filePath);

        // -----------------------------------------------------------------
        // 3. Retrieve all tables by iterating nodes of type NodeType.Table.
        // -----------------------------------------------------------------
        NodeCollection tableNodes = loadedDoc.GetChildNodes(NodeType.Table, true);

        Console.WriteLine($"Total tables found: {tableNodes.Count}");

        int tableIndex = 0;
        foreach (Node node in tableNodes)
        {
            // Cast the node to a Table.
            Table tbl = (Table)node;

            // Output basic information about each table.
            Console.WriteLine($"Table #{tableIndex}: Rows = {tbl.Rows.Count}, Columns = {tbl.FirstRow?.Cells.Count ?? 0}");
            tableIndex++;
        }

        // The program finishes without waiting for user input.
    }
}
