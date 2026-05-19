using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableInsertRowExample
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
            builder.EndRow();

            // Finish the table construction.
            builder.EndTable();

            // Create a new row that will be inserted between the existing rows.
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

            // Insert the new row after the first row (i.e., at index 1).
            // Table.InsertAfter is the documented way to insert a row at a specific position.
            table.InsertAfter(newRow, table.Rows[0]);

            // Optional validation: the table should now contain three rows.
            if (table.Rows.Count != 3)
                throw new InvalidOperationException("Row insertion failed.");

            // Save the document to verify the result.
            doc.Save("InsertedRow.docx");
        }
    }
}
