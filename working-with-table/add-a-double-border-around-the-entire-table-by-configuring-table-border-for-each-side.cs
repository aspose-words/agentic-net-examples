using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

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
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Define the desired border style.
        LineStyle lineStyle = LineStyle.Double;
        double lineWidth = 2.0; // points
        System.Drawing.Color lineColor = System.Drawing.Color.Black;

        // Apply top border to the first row.
        Row firstRow = table.FirstRow;
        firstRow.RowFormat.Borders[BorderType.Top].LineStyle = lineStyle;
        firstRow.RowFormat.Borders[BorderType.Top].LineWidth = lineWidth;
        firstRow.RowFormat.Borders[BorderType.Top].Color = lineColor;

        // Apply bottom border to the last row.
        Row lastRow = table.LastRow;
        lastRow.RowFormat.Borders[BorderType.Bottom].LineStyle = lineStyle;
        lastRow.RowFormat.Borders[BorderType.Bottom].LineWidth = lineWidth;
        lastRow.RowFormat.Borders[BorderType.Bottom].Color = lineColor;

        // Apply left border to the first cell of each row.
        foreach (Row row in table.Rows)
        {
            Cell firstCell = row.FirstCell;
            firstCell.CellFormat.Borders[BorderType.Left].LineStyle = lineStyle;
            firstCell.CellFormat.Borders[BorderType.Left].LineWidth = lineWidth;
            firstCell.CellFormat.Borders[BorderType.Left].Color = lineColor;
        }

        // Apply right border to the last cell of each row.
        foreach (Row row in table.Rows)
        {
            Cell lastCell = row.LastCell;
            lastCell.CellFormat.Borders[BorderType.Right].LineStyle = lineStyle;
            lastCell.CellFormat.Borders[BorderType.Right].LineWidth = lineWidth;
            lastCell.CellFormat.Borders[BorderType.Right].Color = lineColor;
        }

        // Save the document.
        string outputPath = "TableWithDoubleBorder.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
