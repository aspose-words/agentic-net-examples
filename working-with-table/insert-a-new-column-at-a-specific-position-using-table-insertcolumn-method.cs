using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableInsertColumn
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build an initial 2x3 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.InsertCell();
            builder.Write("R1C3");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.InsertCell();
            builder.Write("R2C3");
            builder.EndRow();

            // Finish the table construction.
            builder.EndTable();

            // Insert a new column at position 1 (between the first and second columns).
            int insertPosition = 1;

            // For each row, create a new empty cell and insert it at the desired index.
            foreach (Row row in table.Rows)
            {
                // Create a new cell with an empty paragraph.
                Cell newCell = new Cell(doc);
                newCell.AppendChild(new Paragraph(doc));

                // Insert the new cell into the row's cell collection.
                row.Cells.Insert(insertPosition, newCell);
            }

            // Populate the newly inserted column with sample text.
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                Row row = table.Rows[rowIndex];
                Cell newCell = row.Cells[insertPosition];
                // Add text to the cell's first paragraph.
                newCell.FirstParagraph.AppendChild(new Run(doc, $"R{rowIndex + 1}C{insertPosition + 1}"));
            }

            // Define the output path relative to the executable directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "InsertColumn.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to: {outputPath}");
            }
            else
            {
                throw new InvalidOperationException("Failed to save the document.");
            }
        }
    }
}
