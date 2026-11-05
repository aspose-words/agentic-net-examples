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

        // Define table dimensions.
        int rows = 5;
        int columns = 3;

        // Build the table.
        builder.StartTable();
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                builder.InsertCell();
                builder.Writeln($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Apply distinct border style to the first row.
        Row firstRow = table.FirstRow;
        firstRow.RowFormat.Borders.Top.LineStyle = LineStyle.Single;
        firstRow.RowFormat.Borders.Top.Color = Color.Red;
        firstRow.RowFormat.Borders.Top.LineWidth = 2.0;
        firstRow.RowFormat.Borders.Bottom.LineStyle = LineStyle.Single;
        firstRow.RowFormat.Borders.Bottom.Color = Color.Red;
        firstRow.RowFormat.Borders.Bottom.LineWidth = 2.0;
        firstRow.RowFormat.Borders.Left.LineStyle = LineStyle.Single;
        firstRow.RowFormat.Borders.Left.Color = Color.Red;
        firstRow.RowFormat.Borders.Left.LineWidth = 2.0;
        firstRow.RowFormat.Borders.Right.LineStyle = LineStyle.Single;
        firstRow.RowFormat.Borders.Right.Color = Color.Red;
        firstRow.RowFormat.Borders.Right.LineWidth = 2.0;

        // Apply distinct border style to the last row.
        Row lastRow = table.LastRow;
        lastRow.RowFormat.Borders.Top.LineStyle = LineStyle.Double;
        lastRow.RowFormat.Borders.Top.Color = Color.Blue;
        lastRow.RowFormat.Borders.Top.LineWidth = 2.0;
        lastRow.RowFormat.Borders.Bottom.LineStyle = LineStyle.Double;
        lastRow.RowFormat.Borders.Bottom.Color = Color.Blue;
        lastRow.RowFormat.Borders.Bottom.LineWidth = 2.0;
        lastRow.RowFormat.Borders.Left.LineStyle = LineStyle.Double;
        lastRow.RowFormat.Borders.Left.Color = Color.Blue;
        lastRow.RowFormat.Borders.Left.LineWidth = 2.0;
        lastRow.RowFormat.Borders.Right.LineStyle = LineStyle.Double;
        lastRow.RowFormat.Borders.Right.Color = Color.Blue;
        lastRow.RowFormat.Borders.Right.LineWidth = 2.0;

        // Apply a different border style to all inner cells (rows 2..rows-1).
        for (int r = 1; r < rows - 1; r++)
        {
            Row row = table.Rows[r];
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Borders.Top.LineStyle = LineStyle.Dot;
                cell.CellFormat.Borders.Top.Color = Color.Green;
                cell.CellFormat.Borders.Top.LineWidth = 1.5;

                cell.CellFormat.Borders.Bottom.LineStyle = LineStyle.Dot;
                cell.CellFormat.Borders.Bottom.Color = Color.Green;
                cell.CellFormat.Borders.Bottom.LineWidth = 1.5;

                cell.CellFormat.Borders.Left.LineStyle = LineStyle.Dot;
                cell.CellFormat.Borders.Left.Color = Color.Green;
                cell.CellFormat.Borders.Left.LineWidth = 1.5;

                cell.CellFormat.Borders.Right.LineStyle = LineStyle.Dot;
                cell.CellFormat.Borders.Right.Color = Color.Green;
                cell.CellFormat.Borders.Right.LineWidth = 1.5;
            }
        }

        // Save the document.
        string outputPath = "TableBorders.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
