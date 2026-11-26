using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // First cell.
        builder.InsertCell();
        builder.Writeln("Cell 1");

        // Second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2");

        // End the first row.
        builder.EndRow();

        // Second row - first cell.
        builder.InsertCell();
        builder.Writeln("Cell 3");

        // Second row - second cell.
        builder.InsertCell();
        builder.Writeln("Cell 4");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply custom borders to each cell.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                // Left border: 1 point width.
                cell.CellFormat.Borders[BorderType.Left].LineWidth = 1.0;
                cell.CellFormat.Borders[BorderType.Left].Color = Color.Black;

                // Top border: 2 points width.
                cell.CellFormat.Borders[BorderType.Top].LineWidth = 2.0;
                cell.CellFormat.Borders[BorderType.Top].Color = Color.Black;

                // Right border: 3 points width.
                cell.CellFormat.Borders[BorderType.Right].LineWidth = 3.0;
                cell.CellFormat.Borders[BorderType.Right].Color = Color.Black;

                // Bottom border: 4 points width.
                cell.CellFormat.Borders[BorderType.Bottom].LineWidth = 4.0;
                cell.CellFormat.Borders[BorderType.Bottom].Color = Color.Black;
            }
        }

        // Save the document to disk.
        string outputPath = "CustomCellBorders.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
