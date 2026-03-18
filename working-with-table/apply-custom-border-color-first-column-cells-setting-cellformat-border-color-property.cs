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

        // Build a simple 3‑row, 2‑column table.
        builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Row 1, Col 1");
        builder.InsertCell();
        builder.Write("Row 1, Col 2");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Write("Row 2, Col 1");
        builder.InsertCell();
        builder.Write("Row 2, Col 2");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Apply a custom border color to every cell in the first column.
        foreach (Row row in table.Rows)
        {
            Cell firstCell = row.FirstCell;
            // Setting the Color property of the Borders collection changes all borders of the cell.
            firstCell.CellFormat.Borders.Color = Color.DarkBlue;
        }

        // Save the document.
        doc.Save("FirstColumnBorderColor.docx");
    }
}
