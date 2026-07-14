using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build an initial 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table construction.
        builder.EndTable();

        // Insert a new row at the end of the existing table.
        Row newRow = new Row(doc);
        table.AppendChild(newRow); // Append after the last row.

        // Determine the number of columns from the first row.
        int columnCount = table.FirstRow.Cells.Count;

        // Fill each cell of the new row with placeholder text.
        for (int i = 0; i < columnCount; i++)
        {
            Cell cell = new Cell(doc);
            cell.AppendChild(new Paragraph(doc));
            cell.FirstParagraph.AppendChild(new Run(doc, $"Placeholder {i + 1}"));
            newRow.AppendChild(cell);
        }

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithNewRow.docx");
        doc.Save(outputPath);
    }
}
