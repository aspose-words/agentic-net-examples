using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with 3 rows and 2 columns.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("R1C1");
        builder.InsertCell();
        builder.Writeln("R1C2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("R2C1");
        builder.InsertCell();
        builder.Writeln("R2C2");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Writeln("R3C1");
        builder.InsertCell();
        builder.Writeln("R3C2");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Get the number of rows and columns.
        int rowCount = table.Rows.Count;
        int columnCount = table.Rows[0].Cells.Count;

        // Output the counts to the console.
        Console.WriteLine($"Rows: {rowCount}");
        Console.WriteLine($"Columns: {columnCount}");

        // Save the document to a file.
        doc.Save("TableInfo.docx");
    }
}
