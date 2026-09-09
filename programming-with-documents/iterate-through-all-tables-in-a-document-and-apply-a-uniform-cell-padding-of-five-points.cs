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

        // Build a simple 2x2 table as sample content.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Iterate through all tables in the document.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        foreach (Table tbl in tables)
        {
            // Iterate through each row in the current table.
            foreach (Row row in tbl.Rows)
            {
                // Iterate through each cell in the current row.
                foreach (Cell cell in row.Cells)
                {
                    // Apply uniform padding of 5 points on all sides.
                    cell.CellFormat.SetPaddings(5, 5, 5, 5);
                }
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
