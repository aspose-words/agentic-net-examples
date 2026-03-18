using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class TableRowConditionalBackground
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with two columns: Item and Status.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Status");
        builder.EndRow();

        // Data rows.
        AddDataRow(builder, "Task 1", "Completed");
        AddDataRow(builder, "Task 2", "Pending");
        AddDataRow(builder, "Task 3", "In Progress");
        AddDataRow(builder, "Task 4", "Completed");

        builder.EndTable();

        // Apply background colors to rows based on the value in the Status column.
        // The Status column is the second cell (index 1) in each row.
        foreach (Row row in table.Rows)
        {
            // Skip the header row.
            if (row.IsFirstRow) continue;

            // Retrieve the status text and trim whitespace and end-of-cell markers.
            string status = row.Cells[1].GetText().Trim();

            // Conditional expression to decide the background color.
            Color bgColor = status switch
            {
                "Completed"   => Color.LightGreen,
                "Pending"     => Color.LightYellow,
                "In Progress" => Color.LightBlue,
                _             => Color.Empty // No shading for unknown statuses.
            };

            // Apply the background color to every cell in the current row.
            if (bgColor != Color.Empty)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.Shading.BackgroundPatternColor = bgColor;
                }
            }
        }

        // Save the document.
        doc.Save("TableWithConditionalRowColors.docx");
    }

    // Helper method to insert a data row into the table.
    private static void AddDataRow(DocumentBuilder builder, string item, string status)
    {
        builder.InsertCell();
        builder.Write(item);
        builder.InsertCell();
        builder.Write(status);
        builder.EndRow();
    }
}
