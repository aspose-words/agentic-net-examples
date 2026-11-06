using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to construct a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 3‑row, 2‑column table.
        builder.StartTable();

        // Row 1
        builder.InsertCell();
        builder.Writeln("R1C1");
        builder.InsertCell();
        builder.Writeln("R1C2");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Writeln("R2C1");
        builder.InsertCell();
        builder.Writeln("R2C2");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.Writeln("R3C1");
        builder.InsertCell();
        builder.Writeln("R3C2");
        builder.EndRow();

        builder.EndTable();

        // Save the document (required by the workflow rules).
        doc.Save("TableSample.docx");

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("No table found in the document.");

        // Get row and column counts.
        int rowCount = table.Rows.Count;
        int columnCount = table.Rows[0].Cells.Count;

        // Output the counts.
        Console.WriteLine($"Rows: {rowCount}");
        Console.WriteLine($"Columns: {columnCount}");
    }
}
