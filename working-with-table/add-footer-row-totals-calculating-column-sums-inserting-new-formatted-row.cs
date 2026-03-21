using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class AddFooterRowWithTotals
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table with numeric data.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data rows.
        for (int i = 1; i <= 5; i++)
        {
            builder.InsertCell();
            builder.Write($"Item {i}");
            builder.InsertCell();
            builder.Write((i * 10).ToString()); // numeric value
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // ------------------------------------------------------------
        // Calculate column sums (skip the header row, assume numeric data
        // is in the second column).
        // ------------------------------------------------------------
        int columnCount = table.FirstRow.Cells.Count;
        double[] columnSums = new double[columnCount];

        // Start from row index 1 to skip header.
        for (int rowIdx = 1; rowIdx < table.Rows.Count; rowIdx++)
        {
            Row row = table.Rows[rowIdx];
            for (int colIdx = 0; colIdx < columnCount; colIdx++)
            {
                string cellText = row.Cells[colIdx].GetText().Trim();
                // Try to parse a double; if it fails, treat it as zero.
                if (double.TryParse(cellText, out double value))
                    columnSums[colIdx] += value;
            }
        }

        // ------------------------------------------------------------
        // Insert a new footer row with the calculated totals.
        // ------------------------------------------------------------
        // Add a new empty row at the end of the table.
        Row totalRow = new Row(doc);
        table.Rows.Add(totalRow);

        // Ensure the row has the correct number of cells.
        for (int i = 0; i < columnCount; i++)
            totalRow.Cells.Add(new Cell(doc));

        // Apply formatting to the total row (bold text and a top border).
        totalRow.RowFormat.Height = 20;
        totalRow.RowFormat.HeightRule = HeightRule.AtLeast;
        totalRow.RowFormat.Borders[BorderType.Top].LineStyle = LineStyle.Single;
        totalRow.RowFormat.Borders[BorderType.Top].Color = Color.Black;
        builder.Font.Bold = true;

        // Fill cells of the total row.
        for (int colIdx = 0; colIdx < columnCount; colIdx++)
        {
            Cell cell = totalRow.Cells[colIdx];
            cell.EnsureMinimum();

            string text = colIdx == 0 ? "Total" : columnSums[colIdx].ToString();
            cell.FirstParagraph.Runs.Clear(); // Remove any existing runs.
            cell.FirstParagraph.AppendChild(new Run(doc, text));
        }

        // Reset bold formatting for subsequent operations.
        builder.Font.Bold = false;

        // Save the document.
        doc.Save("TableWithTotals.docx");
    }
}
