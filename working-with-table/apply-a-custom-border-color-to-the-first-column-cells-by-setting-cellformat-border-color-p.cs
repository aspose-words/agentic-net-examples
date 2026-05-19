using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3‑row, 2‑column table.
        Table table = builder.StartTable();

        for (int row = 0; row < 3; row++)
        {
            // First column cell.
            builder.InsertCell();
            builder.Write($"Row {row + 1}, Col 1");

            // Second column cell.
            builder.InsertCell();
            builder.Write($"Row {row + 1}, Col 2");

            builder.EndRow();
        }

        builder.EndTable();

        // Apply a custom border color (Blue) to every cell in the first column.
        foreach (Row r in table.Rows)
        {
            Cell firstCell = r.FirstCell;

            // Set each side of the cell border to the desired color.
            firstCell.CellFormat.Borders[BorderType.Left].Color = Color.Blue;
            firstCell.CellFormat.Borders[BorderType.Right].Color = Color.Blue;
            firstCell.CellFormat.Borders[BorderType.Top].Color = Color.Blue;
            firstCell.CellFormat.Borders[BorderType.Bottom].Color = Color.Blue;
        }

        // Save the document to the local file system.
        doc.Save("FirstColumnBorderColor.docx");
    }
}
