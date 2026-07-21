using System;
using System.IO;
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
            builder.StartTable();

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
            builder.EndTable();

            // Retrieve the created table.
            Table table = doc.FirstSection.Body.Tables[0];

            // Create a new row that will be inserted.
            Row newRow = new Row(doc);

            // Add two cells to the new row (matching the existing column count).
            Cell cell1 = new Cell(doc);
            cell1.AppendChild(new Paragraph(doc));
            cell1.FirstParagraph.AppendChild(new Run(doc, "Inserted Row, Cell 1"));
            newRow.AppendChild(cell1);

            Cell cell2 = new Cell(doc);
            cell2.AppendChild(new Paragraph(doc));
            cell2.FirstParagraph.AppendChild(new Run(doc, "Inserted Row, Cell 2"));
            newRow.AppendChild(cell2);

            // Insert the new row at index 1 (between the original rows).
            // RowCollection.Insert inserts the row at the specified position.
            table.Rows.Insert(1, newRow);

            // Optional validation: the table should now contain three rows.
            if (table.Rows.Count != 3)
                throw new InvalidOperationException("Row insertion failed.");

            // Save the document to the local file system.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "InsertedRow.docx");
            doc.Save(outputPath);

            // Indicate successful completion.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
