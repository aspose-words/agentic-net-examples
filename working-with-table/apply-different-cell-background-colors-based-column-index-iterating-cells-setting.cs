using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 4‑row by 3‑column table.
        builder.StartTable();
        for (int r = 0; r < 4; r++)
        {
            for (int c = 0; c < 3; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Iterate over each cell and set its background color based on column index.
        foreach (Row row in table.Rows)
        {
            for (int colIndex = 0; colIndex < row.Cells.Count; colIndex++)
            {
                Cell cell = row.Cells[colIndex];

                // Even columns get LightBlue, odd columns get LightGray.
                if (colIndex % 2 == 0)
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
                else
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            }
        }

        // Save the document.
        doc.Save("ColoredTable.docx");
    }
}
