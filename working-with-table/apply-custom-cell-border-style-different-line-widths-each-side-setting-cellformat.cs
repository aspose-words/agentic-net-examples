using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class CustomCellBorders
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a single cell.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell with custom borders");

        // Retrieve the first cell that was just created.
        Cell cell = table.FirstRow.FirstCell;

        // Set different line widths, styles, and colors for each side of the cell.
        cell.CellFormat.Borders[BorderType.Left].LineWidth = 1.0;               // 1 point left border
        cell.CellFormat.Borders[BorderType.Left].LineStyle = LineStyle.Single;
        cell.CellFormat.Borders[BorderType.Left].Color = Color.Red;

        cell.CellFormat.Borders[BorderType.Right].LineWidth = 2.5;              // 2.5 points right border
        cell.CellFormat.Borders[BorderType.Right].LineStyle = LineStyle.Double;
        cell.CellFormat.Borders[BorderType.Right].Color = Color.Green;

        cell.CellFormat.Borders[BorderType.Top].LineWidth = 0.5;                // 0.5 point top border
        // "Dash" is not a valid member of the LineStyle enum. Use a dash‑style value such as DashSmallGap.
        cell.CellFormat.Borders[BorderType.Top].LineStyle = LineStyle.DashSmallGap;
        cell.CellFormat.Borders[BorderType.Top].Color = Color.Blue;

        cell.CellFormat.Borders[BorderType.Bottom].LineWidth = 3.0;             // 3 points bottom border
        cell.CellFormat.Borders[BorderType.Bottom].LineStyle = LineStyle.Thick;
        cell.CellFormat.Borders[BorderType.Bottom].Color = Color.Purple;

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "CustomCellBorders.docx";
        doc.Save(outputPath);
    }
}
