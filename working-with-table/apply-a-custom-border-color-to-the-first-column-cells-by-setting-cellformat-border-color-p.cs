using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 5x3 table.
        builder.StartTable();
        for (int i = 0; i < 5; i++) // rows
        {
            for (int j = 0; j < 3; j++) // columns
            {
                builder.InsertCell();
                builder.Writeln($"Row {i + 1}, Col {j + 1}");
                // End the row after the last cell.
                if (j == 2)
                    builder.EndRow();
            }
        }
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Apply a custom border color to the cells in the first column.
        foreach (Row row in table.Rows)
        {
            Cell firstCell = row.Cells[0];
            // Set the color for each side of the cell border.
            firstCell.CellFormat.Borders[BorderType.Left].Color = Color.Red;
            firstCell.CellFormat.Borders[BorderType.Right].Color = Color.Red;
            firstCell.CellFormat.Borders[BorderType.Top].Color = Color.Red;
            firstCell.CellFormat.Borders[BorderType.Bottom].Color = Color.Red;
        }

        // Save the document to disk.
        string outputPath = "FirstColumnBorderColor.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
