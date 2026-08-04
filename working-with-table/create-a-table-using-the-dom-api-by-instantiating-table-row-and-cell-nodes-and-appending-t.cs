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

        // Create a new table and add it to the document's first section body.
        Table table = new Table(doc);
        doc.FirstSection.Body.AppendChild(table);

        // Build a 2x2 table using the DOM API.
        for (int rowIdx = 0; rowIdx < 2; rowIdx++)
        {
            // Create a new row and append it to the table.
            Row row = new Row(doc);
            table.AppendChild(row);

            for (int colIdx = 0; colIdx < 2; colIdx++)
            {
                // Create a new cell.
                Cell cell = new Cell(doc);

                // Each cell must contain at least one paragraph.
                Paragraph paragraph = new Paragraph(doc);
                cell.AppendChild(paragraph);

                // Add text to the paragraph.
                Run run = new Run(doc, $"Row {rowIdx + 1}, Cell {colIdx + 1}");
                paragraph.AppendChild(run);

                // Append the cell to the current row.
                row.AppendChild(cell);
            }
        }

        // Save the document to a file in the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CreatedTable.docx");
        doc.Save(outputPath);

        // Indicate that the document has been saved.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
