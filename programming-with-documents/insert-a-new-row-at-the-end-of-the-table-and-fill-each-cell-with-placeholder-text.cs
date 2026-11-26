using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with 2 rows and 3 columns as initial content.
        Table table = builder.StartTable();

        // First row.
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Write($"Row1Col{col + 1}");
        }
        builder.EndRow();

        // Second row.
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Write($"Row2Col{col + 1}");
        }
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Insert a new row at the end of the existing table.
        // The Rows collection is zero‑based; the new row will be added after the last existing row.
        Row newRow = new Row(doc);
        // Ensure the new row has the same number of cells as the existing rows.
        int columnCount = table.Rows[0].Cells.Count;
        for (int i = 0; i < columnCount; i++)
        {
            // Add a new cell to the row.
            Cell cell = new Cell(doc);
            // Insert placeholder text into the cell.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, "Placeholder");
            para.AppendChild(run);
            cell.AppendChild(para);
            newRow.AppendChild(cell);
        }
        // Append the new row to the table.
        table.Rows.Add(newRow);

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
