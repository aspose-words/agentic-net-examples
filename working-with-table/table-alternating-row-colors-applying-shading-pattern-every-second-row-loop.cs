using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class AlternatingRowColors
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        int rowCount = 10;   // Total number of rows.
        int colCount = 3;    // Number of columns per row.

        for (int i = 0; i < rowCount; i++)
        {
            // Insert cells for the current row.
            for (int j = 0; j < colCount; j++)
            {
                builder.InsertCell();
                builder.Write($"R{i + 1}C{j + 1}");
            }

            // End the current row.
            builder.EndRow();

            // Apply shading to every second (even-indexed) row.
            // i is zero‑based, so rows 2,4,6,... have i % 2 == 1.
            if (i % 2 == 1)
            {
                Row currentRow = table.LastRow;
                foreach (Cell cell in currentRow.Cells)
                {
                    // Set the background color for the cell.
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
                }
            }
        }

        // Finish the table.
        builder.EndTable();

        // Save the document.
        doc.Save("AlternatingRows.docx");
    }
}
