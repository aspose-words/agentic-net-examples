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

        // Build a simple 2x2 table.
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

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Apply thin inner borders to all cells.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Borders[BorderType.Left].LineWidth = 0.5;
                cell.CellFormat.Borders[BorderType.Right].LineWidth = 0.5;
                cell.CellFormat.Borders[BorderType.Top].LineWidth = 0.5;
                cell.CellFormat.Borders[BorderType.Bottom].LineWidth = 0.5;

                cell.CellFormat.Borders[BorderType.Left].Color = Color.Black;
                cell.CellFormat.Borders[BorderType.Right].Color = Color.Black;
                cell.CellFormat.Borders[BorderType.Top].Color = Color.Black;
                cell.CellFormat.Borders[BorderType.Bottom].Color = Color.Black;
            }
        }

        // Apply thick outer borders.
        int lastRowIndex = table.Rows.Count - 1;
        int lastCellIndex = table.Rows[0].Cells.Count - 1;

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            Row row = table.Rows[rowIndex];
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                Cell cell = row.Cells[cellIndex];

                // Top border for first row.
                if (rowIndex == 0)
                {
                    cell.CellFormat.Borders[BorderType.Top].LineWidth = 2.0;
                }

                // Bottom border for last row.
                if (rowIndex == lastRowIndex)
                {
                    cell.CellFormat.Borders[BorderType.Bottom].LineWidth = 2.0;
                }

                // Left border for first column.
                if (cellIndex == 0)
                {
                    cell.CellFormat.Borders[BorderType.Left].LineWidth = 2.0;
                }

                // Right border for last column.
                if (cellIndex == lastCellIndex)
                {
                    cell.CellFormat.Borders[BorderType.Right].LineWidth = 2.0;
                }
            }
        }

        // Save the document.
        string outputPath = "TableStyle.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
