using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for constructing content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table (2 rows x 2 columns) to demonstrate padding changes.
        builder.StartTable();

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Iterate through all tables in the document.
        foreach (Table tbl in doc.GetChildNodes(NodeType.Table, true))
        {
            // Iterate through each row in the current table.
            foreach (Row row in tbl.Rows)
            {
                // Iterate through each cell in the current row.
                foreach (Cell cell in row.Cells)
                {
                    // Apply a uniform padding of five points to all sides of the cell.
                    cell.CellFormat.TopPadding = 5;
                    cell.CellFormat.BottomPadding = 5;
                    cell.CellFormat.LeftPadding = 5;
                    cell.CellFormat.RightPadding = 5;
                }
            }
        }

        // Save the modified document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
