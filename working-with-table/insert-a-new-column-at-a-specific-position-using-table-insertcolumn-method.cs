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

            // Finish the table and keep a reference to it.
            Table table = builder.EndTable();

            // Insert a new column at index 1 (between the original first and second columns).
            int insertIndex = 1; // zero‑based index where the new column will appear
            foreach (Row row in table.Rows)
            {
                // Create a new cell with an empty paragraph.
                Cell newCell = new Cell(doc);
                newCell.AppendChild(new Paragraph(doc));

                // Insert the cell at the desired position.
                if (insertIndex < row.Cells.Count)
                {
                    // Insert before the cell that currently occupies the target index.
                    Cell referenceCell = row.Cells[insertIndex];
                    row.InsertBefore(newCell, referenceCell);
                }
                else
                {
                    // If the index is beyond the current count, simply append.
                    row.AppendChild(newCell);
                }

                // Populate the newly inserted cell.
                newCell.FirstParagraph.AppendChild(new Run(doc, "New"));
            }

            // Validate that each row now contains four cells.
            foreach (Row row in table.Rows)
            {
                if (row.Cells.Count != 4)
                {
                    throw new InvalidOperationException("Column insertion failed: each row must have 4 cells.");
                }
            }

            // Save the document to the local file system.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "InsertColumn.docx");
            doc.Save(outputPath);
        }
    }
}
