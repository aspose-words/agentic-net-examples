using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class TableBorderExample
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // ---- Row 1 (First row) ----
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // ---- Row 2 (Inner row) ----
        builder.InsertCell();
        builder.Write("Item A1");
        builder.InsertCell();
        builder.Write("Item A2");
        builder.EndRow();

        // ---- Row 3 (Inner row) ----
        builder.InsertCell();
        builder.Write("Item B1");
        builder.InsertCell();
        builder.Write("Item B2");
        builder.EndRow();

        // ---- Row 4 (Last row) ----
        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();
        builder.Write("Sum");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Apply distinct border styles to the first row.
        Row firstRow = table.FirstRow;
        firstRow.RowFormat.Borders.LineStyle = LineStyle.Double;
        firstRow.RowFormat.Borders.Color = Color.Blue;
        firstRow.RowFormat.Borders.LineWidth = 2.0;

        // Apply distinct border styles to the last row.
        Row lastRow = table.LastRow;
        // "DashDot" is not a valid LineStyle in Aspose.Words. Use a supported style such as DashSmallGap.
        lastRow.RowFormat.Borders.LineStyle = LineStyle.DashSmallGap;
        lastRow.RowFormat.Borders.Color = Color.Red;
        lastRow.RowFormat.Borders.LineWidth = 2.0;

        // Apply a uniform border style to all inner rows.
        for (int i = 1; i < table.Rows.Count - 1; i++)
        {
            Row innerRow = table.Rows[i];
            innerRow.RowFormat.Borders.LineStyle = LineStyle.Single;
            innerRow.RowFormat.Borders.Color = Color.Green;
            innerRow.RowFormat.Borders.LineWidth = 1.0;
        }

        // Optionally, set cell borders for inner cells to demonstrate CellFormat usage.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                // Skip first and last rows (already styled via RowFormat).
                if (row == firstRow || row == lastRow) continue;

                // Set a thin gray border around each inner cell.
                cell.CellFormat.Borders.LineStyle = LineStyle.Single;
                cell.CellFormat.Borders.Color = Color.Gray;
                cell.CellFormat.Borders.LineWidth = 0.5;
            }
        }

        // Save the document.
        doc.Save("TableWithDistinctBorders.docx");
    }
}
