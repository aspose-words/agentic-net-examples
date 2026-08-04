using System;
using System.IO;
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

            // Initialize DocumentBuilder for building the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build an initial 2x2 table.
            builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1, Row 1");
            builder.InsertCell();
            builder.Write("Cell 2, Row 1");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 1, Row 2");
            builder.InsertCell();
            builder.Write("Cell 2, Row 2");
            builder.EndRow();

            // Finish the table and obtain a reference to it.
            Table table = builder.EndTable();

            // ------------------------------------------------------------
            // Insert a new row at the end of the existing table.
            // ------------------------------------------------------------

            // Create a new row associated with the document.
            Row newRow = new Row(doc);

            // Append the new row to the table.
            table.AppendChild(newRow);

            // Define the number of columns (matching the existing table).
            int columnCount = 2;

            // Populate each cell in the new row with placeholder text.
            for (int i = 0; i < columnCount; i++)
            {
                // Create a new cell.
                Cell cell = new Cell(doc);

                // Ensure the cell contains a paragraph.
                cell.AppendChild(new Paragraph(doc));

                // Add placeholder text to the paragraph.
                cell.FirstParagraph.AppendChild(new Run(doc, $"Placeholder {i + 1}"));

                // Append the cell to the new row.
                newRow.AppendChild(cell);
            }

            // Save the document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "InsertedRow.docx");
            doc.Save(outputPath);
        }
    }
}
