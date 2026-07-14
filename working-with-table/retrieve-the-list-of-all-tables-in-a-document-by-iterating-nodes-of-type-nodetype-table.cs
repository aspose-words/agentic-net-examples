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

        // Build first table.
        Table table1 = builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 1 - Cell 1");
        builder.InsertCell();
        builder.Write("Table 1 - Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Build second table.
        Table table2 = builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 2 - Cell 1");
        builder.InsertCell();
        builder.Write("Table 2 - Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Save the sample document (optional, demonstrates file creation).
        doc.Save("Sample.docx");

        // Retrieve all tables in the document by iterating nodes of type NodeType.Table.
        NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);

        // Output information about each table.
        for (int i = 0; i < tableNodes.Count; i++)
        {
            Table tbl = (Table)tableNodes[i];
            int rowCount = tbl.Rows.Count;
            int columnCount = tbl.FirstRow?.Cells.Count ?? 0;
            Console.WriteLine($"Table {i}: {rowCount} rows, {columnCount} columns");

            // Optionally, print the text of each cell.
            for (int r = 0; r < rowCount; r++)
            {
                Row row = tbl.Rows[r];
                for (int c = 0; c < row.Cells.Count; c++)
                {
                    string cellText = row.Cells[c].ToString(SaveFormat.Text).Trim();
                    Console.WriteLine($"  Cell[{r},{c}] = \"{cellText}\"");
                }
            }
        }
    }
}
