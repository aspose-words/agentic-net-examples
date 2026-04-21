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

        // Create a Table node that belongs to the document.
        Table table = new Table(doc);

        // Append the table to the body of the first section.
        doc.FirstSection.Body.AppendChild(table);

        // Build a 2x2 table using the DOM API (Table, Row, Cell).
        for (int rowIdx = 0; rowIdx < 2; rowIdx++)
        {
            // Create a new Row and add it to the table.
            Row row = new Row(doc);
            table.AppendChild(row);

            for (int colIdx = 0; colIdx < 2; colIdx++)
            {
                // Create a new Cell.
                Cell cell = new Cell(doc);

                // Each cell must contain at least one Paragraph.
                Paragraph paragraph = new Paragraph(doc);
                cell.AppendChild(paragraph);

                // Add text to the paragraph.
                Run run = new Run(doc, $"R{rowIdx + 1}C{colIdx + 1}");
                paragraph.AppendChild(run);

                // Append the cell to the current row.
                row.AppendChild(cell);
            }
        }

        // Save the document to a file in the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableCreated.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not created.");
        }

        // Inform the user where the file was saved (no waiting for input).
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
