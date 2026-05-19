using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a new table and associate it with the document.
            Table table = new Table(doc);

            // Append the table to the document's first section body.
            doc.FirstSection.Body.AppendChild(table);

            // Build a 2x2 table manually using the DOM API.
            for (int rowIndex = 0; rowIndex < 2; rowIndex++)
            {
                // Create a new row.
                Row row = new Row(doc);
                table.AppendChild(row);

                for (int cellIndex = 0; cellIndex < 2; cellIndex++)
                {
                    // Create a new cell.
                    Cell cell = new Cell(doc);
                    row.AppendChild(cell);

                    // Every cell must contain at least one paragraph.
                    Paragraph para = new Paragraph(doc);
                    cell.AppendChild(para);

                    // Add some text to the paragraph.
                    Run run = new Run(doc, $"Row {rowIndex + 1}, Cell {cellIndex + 1}");
                    para.AppendChild(run);
                }
            }

            // Define the output path. Ensure the directory exists.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CreatedTable.docx");
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Table created and saved successfully to: {outputPath}");
            }
            else
            {
                throw new InvalidOperationException("Failed to save the document.");
            }
        }
    }
}
