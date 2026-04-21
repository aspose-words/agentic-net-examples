using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableInsertRow
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build an initial 2‑row, 2‑column table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // -----------------------------------------------------------------
            // Insert a new row at index 1 (between the existing rows) using RowCollection.Insert.
            // -----------------------------------------------------------------
            // Create a new row and populate its cells.
            Row newRow = new Row(doc);
            // Ensure the row has at least one cell before adding content.
            newRow.EnsureMinimum();

            // First cell of the new row.
            Cell cell1 = new Cell(doc);
            cell1.AppendChild(new Paragraph(doc));
            cell1.FirstParagraph.AppendChild(new Run(doc, "Inserted C1"));
            newRow.AppendChild(cell1);

            // Second cell of the new row.
            Cell cell2 = new Cell(doc);
            cell2.AppendChild(new Paragraph(doc));
            cell2.FirstParagraph.AppendChild(new Run(doc, "Inserted C2"));
            newRow.AppendChild(cell2);

            // Insert the row at the desired index.
            // RowCollection.Insert inserts the supplied row at the specified zero‑based index.
            table.Rows.Insert(1, newRow);

            // Simple validation – the table should now contain three rows.
            if (table.Rows.Count != 3)
                throw new InvalidOperationException("Row insertion failed.");

            // Save the document to the local file system.
            string outputPath = "InsertedRow.docx";
            doc.Save(outputPath);

            // Inform that the operation completed.
            Console.WriteLine($"Document saved to '{outputPath}'. Row count: {table.Rows.Count}");
        }
    }
}
