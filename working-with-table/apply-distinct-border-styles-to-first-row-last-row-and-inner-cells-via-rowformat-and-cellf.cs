using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for constructing content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 3x3 table.
        builder.StartTable();
        for (int row = 0; row < 3; row++)
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

        // Apply distinct border style to the first row.
        Row firstRow = table.Rows[0];
        firstRow.RowFormat.Borders.LineStyle = LineStyle.Single;
        firstRow.RowFormat.Borders.Color = Color.Red;
        firstRow.RowFormat.Borders.LineWidth = 2.0; // points

        // Apply distinct border style to the last row.
        Row lastRow = table.Rows[table.Rows.Count - 1];
        lastRow.RowFormat.Borders.LineStyle = LineStyle.Single;
        lastRow.RowFormat.Borders.Color = Color.Blue;
        lastRow.RowFormat.Borders.LineWidth = 2.0; // points

        // Apply a different border style to inner cells (rows that are not first or last).
        for (int i = 1; i < table.Rows.Count - 1; i++)
        {
            Row innerRow = table.Rows[i];
            foreach (Cell cell in innerRow.Cells)
            {
                cell.CellFormat.Borders.LineStyle = LineStyle.Dot; // corrected enum value
                cell.CellFormat.Borders.Color = Color.Green;
                cell.CellFormat.Borders.LineWidth = 1.5; // points
            }
        }

        // Save the document.
        string outputPath = "TableBorders.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
