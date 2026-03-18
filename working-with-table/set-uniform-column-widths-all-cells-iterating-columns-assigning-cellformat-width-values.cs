using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SetUniformColumnWidths
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x2 table.
        Table table = builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.InsertCell();
        builder.Write("R1C3");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.InsertCell();
        builder.Write("R2C3");
        builder.EndRow();

        builder.EndTable();

        // Desired uniform width for each column (in points).
        double uniformWidth = 100.0;

        // Iterate over each column index.
        int columnCount = table.FirstRow.Cells.Count;
        for (int col = 0; col < columnCount; col++)
        {
            // Iterate over each row in the table.
            foreach (Row row in table.Rows)
            {
                // Set the width of the cell in the current column.
                Cell cell = row.Cells[col];
                cell.CellFormat.Width = uniformWidth;
            }
        }

        // Save the document.
        doc.Save("UniformColumnWidths.docx");
    }
}
