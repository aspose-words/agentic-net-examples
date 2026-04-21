using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableInsertColumnExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2‑row, 3‑column table.
            builder.StartTable();

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

            // Finish the table and obtain a reference to it.
            Table table = builder.EndTable();

            // Insert a new column at index 1 (between the original first and second columns).
            int insertIndex = 1;
            foreach (Row row in table.Rows)
            {
                // Create a new empty cell with an empty paragraph.
                Cell newCell = new Cell(doc);
                newCell.AppendChild(new Paragraph(doc));

                // Insert the cell at the desired position within the row.
                row.Cells.Insert(insertIndex, newCell);
            }

            // Populate the newly inserted column with sample text.
            int rowNumber = 1;
            foreach (Row row in table.Rows)
            {
                Cell cell = row.Cells[insertIndex];
                cell.FirstParagraph.AppendChild(new Run(doc, $"R{rowNumber}CNew"));
                rowNumber++;
            }

            // Save the document to the local file system.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "InsertColumn.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
