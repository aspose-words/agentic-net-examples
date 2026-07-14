using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a table and add it to the document body.
        Table table = new Table(doc);
        doc.FirstSection.Body.AppendChild(table);

        // Build a 2x2 table using the DOM API.
        for (int rowIndex = 0; rowIndex < 2; rowIndex++)
        {
            // Create a new row and add it to the table.
            Row row = new Row(doc);
            table.AppendChild(row);

            for (int colIndex = 0; colIndex < 2; colIndex++)
            {
                // Create a new cell.
                Cell cell = new Cell(doc);

                // Each cell must contain at least one paragraph.
                Paragraph paragraph = new Paragraph(doc);
                cell.AppendChild(paragraph);

                // Add some text to the paragraph.
                Run run = new Run(doc, $"Row {rowIndex + 1}, Cell {colIndex + 1}");
                paragraph.AppendChild(run);

                // Add the cell to the current row.
                row.AppendChild(cell);
            }
        }

        // Save the document to a file.
        doc.Save("TableFromDom.docx");
    }
}
