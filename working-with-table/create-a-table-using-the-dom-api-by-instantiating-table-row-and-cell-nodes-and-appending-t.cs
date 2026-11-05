using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a new table node.
        Table table = new Table(doc);

        // Build a 2x3 table (2 rows, 3 columns).
        for (int rowIndex = 0; rowIndex < 2; rowIndex++)
        {
            // Create a new row.
            Row row = new Row(doc);

            for (int colIndex = 0; colIndex < 3; colIndex++)
            {
                // Create a new cell.
                Cell cell = new Cell(doc);

                // Add a paragraph with some text to the cell.
                Paragraph para = new Paragraph(doc);
                Run run = new Run(doc, $"Row {rowIndex + 1}, Cell {colIndex + 1}");
                para.AppendChild(run);
                cell.AppendChild(para);

                // Append the cell to the row.
                row.AppendChild(cell);
            }

            // Append the row to the table.
            table.AppendChild(row);
        }

        // Append the table to the document's first section body.
        doc.FirstSection.Body.AppendChild(table);

        // Define output file path.
        string outputPath = "TableCreated.docx";

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
