using System;
using System.IO;
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

        // Build a table with 4 rows and 3 columns.
        builder.StartTable();
        for (int row = 0; row < 4; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Apply distinct border styles to the first row.
        Row firstRow = table.FirstRow;
        firstRow.RowFormat.Borders[BorderType.Top].LineStyle = LineStyle.Single;
        firstRow.RowFormat.Borders[BorderType.Top].LineWidth = 2.0; // points
        firstRow.RowFormat.Borders[BorderType.Top].Color = Color.Red;
        firstRow.RowFormat.Borders[BorderType.Bottom].LineStyle = LineStyle.Double;
        firstRow.RowFormat.Borders[BorderType.Bottom].LineWidth = 1.5;
        firstRow.RowFormat.Borders[BorderType.Bottom].Color = Color.Blue;

        // Apply distinct border styles to the last row.
        Row lastRow = table.LastRow;
        lastRow.RowFormat.Borders[BorderType.Top].LineStyle = LineStyle.Double;
        lastRow.RowFormat.Borders[BorderType.Top].LineWidth = 1.5;
        lastRow.RowFormat.Borders[BorderType.Top].Color = Color.Purple;
        lastRow.RowFormat.Borders[BorderType.Bottom].LineStyle = LineStyle.Single;
        lastRow.RowFormat.Borders[BorderType.Bottom].LineWidth = 2.0;
        lastRow.RowFormat.Borders[BorderType.Bottom].Color = Color.Green;

        // Apply thin gray borders to inner cells (excluding first and last rows).
        for (int i = 1; i < table.Rows.Count - 1; i++)
        {
            Row innerRow = table.Rows[i];
            foreach (Cell cell in innerRow.Cells)
            {
                cell.CellFormat.Borders[BorderType.Left].LineStyle = LineStyle.Single;
                cell.CellFormat.Borders[BorderType.Left].LineWidth = 0.5;
                cell.CellFormat.Borders[BorderType.Left].Color = Color.Gray;

                cell.CellFormat.Borders[BorderType.Right].LineStyle = LineStyle.Single;
                cell.CellFormat.Borders[BorderType.Right].LineWidth = 0.5;
                cell.CellFormat.Borders[BorderType.Right].Color = Color.Gray;
            }
        }

        // Save the document.
        string outputPath = "TableBorders.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
