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

        // Build an initial table with 2 rows and 3 columns.
        builder.StartTable();

        // First row
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Write($"R1C{col + 1}");
        }
        builder.EndRow();

        // Second row
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Write($"R2C{col + 1}");
        }
        builder.EndRow();

        // Finish the table.
        Table table = builder.EndTable();

        // Insert a new row at the end of the existing table.
        Row newRow = new Row(doc);
        // Ensure the new row has the same number of cells as the existing rows.
        int cellCount = table.FirstRow.Cells.Count;
        for (int i = 0; i < cellCount; i++)
        {
            Cell cell = new Cell(doc);
            // Each cell needs at least one paragraph.
            cell.AppendChild(new Paragraph(doc));
            // Add placeholder text to the cell.
            cell.FirstParagraph.AppendChild(new Run(doc, "Placeholder"));
            newRow.AppendChild(cell);
        }

        // Append the new row to the table.
        table.AppendChild(newRow);

        // Save the document to the local file system.
        doc.Save("TableWithAddedRow.docx");
    }
}
