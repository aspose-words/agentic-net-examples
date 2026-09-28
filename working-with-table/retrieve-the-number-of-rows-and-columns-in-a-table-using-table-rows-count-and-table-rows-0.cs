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

        // Build a table with 2 rows and 3 columns.
        builder.StartTable();

        // First row.
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Writeln($"R1C{col + 1}");
        }
        builder.EndRow();

        // Second row.
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Writeln($"R2C{col + 1}");
        }
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Get row and column counts.
        int rowCount = table.Rows.Count;
        int columnCount = table.Rows[0].Cells.Count;

        // Output the counts.
        Console.WriteLine($"Rows: {rowCount}");
        Console.WriteLine($"Columns: {columnCount}");

        // Save the document.
        doc.Save("SampleTable.docx");
    }
}
