using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample data for the table.
        string[] headers = { "ID", "Name", "Description" };
        string[,] rows = {
            { "1", "Alice", "Software engineer with 5 years experience" },
            { "2", "Bob", "Data analyst" },
            { "3", "Charlie", "Project manager overseeing multiple teams" }
        };

        // Build the table.
        builder.StartTable();

        // Insert header row.
        foreach (string header in headers)
        {
            builder.InsertCell();
            builder.Font.Bold = true;
            builder.Writeln(header);
        }
        builder.EndRow();

        // Insert data rows.
        for (int r = 0; r < rows.GetLength(0); r++)
        {
            for (int c = 0; c < rows.GetLength(1); c++)
            {
                builder.InsertCell();
                builder.Font.Bold = false;
                builder.Writeln(rows[r, c]);
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Simple heuristic to estimate text width (points) based on character count and font size.
        double MeasureTextWidth(string text, double fontSize)
        {
            // Approximate each character as 0.5 * fontSize points.
            return text.Length * fontSize * 0.5;
        }

        // Determine maximum required width for each column.
        int columnCount = table.Rows[0].Cells.Count;
        double[] maxColumnWidths = new double[columnCount];

        // Include header row in measurement.
        for (int col = 0; col < columnCount; col++)
        {
            Cell headerCell = table.Rows[0].Cells[col];
            string headerText = headerCell.GetText().TrimEnd('\a'); // Remove cell end marker.
            double headerWidth = MeasureTextWidth(headerText, builder.Font.Size);
            maxColumnWidths[col] = headerWidth;
        }

        // Measure data rows.
        for (int rowIdx = 1; rowIdx < table.Rows.Count; rowIdx++)
        {
            Row row = table.Rows[rowIdx];
            for (int col = 0; col < columnCount; col++)
            {
                Cell cell = row.Cells[col];
                string cellText = cell.GetText().TrimEnd('\a');
                double cellWidth = MeasureTextWidth(cellText, builder.Font.Size);
                if (cellWidth > maxColumnWidths[col])
                    maxColumnWidths[col] = cellWidth;
            }
        }

        // Apply calculated widths to each column (add a small margin).
        const double marginPoints = 5.0;
        for (int col = 0; col < columnCount; col++)
        {
            double finalWidth = maxColumnWidths[col] + marginPoints;
            // Set the width for every cell in this column.
            foreach (Row row in table.Rows)
            {
                Cell cell = row.Cells[col];
                cell.CellFormat.Width = finalWidth;
            }
        }

        // Fix column widths by using FixedColumnWidths behavior.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Save the document.
        string outputPath = "TableColumnWidths.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
