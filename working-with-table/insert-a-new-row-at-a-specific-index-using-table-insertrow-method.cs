using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsInsertRowExample
{
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
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndTable(); // Ends the table construction.

            // At this point the table has two rows. We will insert a new row at index 1 (between the existing rows).

            // Create a new row that matches the table's column count.
            Row newRow = new Row(doc);

            // First cell of the new row.
            Cell cell1 = new Cell(doc);
            cell1.AppendChild(new Paragraph(doc));
            cell1.FirstParagraph.AppendChild(new Run(doc, "Inserted Row, Cell 1"));
            newRow.AppendChild(cell1);

            // Second cell of the new row.
            Cell cell2 = new Cell(doc);
            cell2.AppendChild(new Paragraph(doc));
            cell2.FirstParagraph.AppendChild(new Run(doc, "Inserted Row, Cell 2"));
            newRow.AppendChild(cell2);

            // Insert the new row at the desired index using the RowCollection.Insert method.
            // The index is zero‑based; 1 means after the first row.
            table.Rows.Insert(1, newRow);

            // Simple validation to ensure the row was inserted.
            if (table.Rows.Count != 3)
                throw new InvalidOperationException("Row insertion failed; expected 3 rows.");

            // Save the document to the local file system.
            string outputPath = "InsertRowExample.docx";
            doc.Save(outputPath);

            // Inform that the process completed successfully.
            Console.WriteLine($"Document saved to '{outputPath}'. Table now contains {table.Rows.Count} rows.");
        }
    }
}
