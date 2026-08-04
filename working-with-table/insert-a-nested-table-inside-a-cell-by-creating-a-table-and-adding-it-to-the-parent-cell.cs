using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace NestedTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the outer table (2 columns, 1 row for simplicity).
            Table outerTable = builder.StartTable();

            // First cell of the outer table.
            builder.InsertCell();
            builder.Write("Outer Cell 1");

            // Second cell of the outer table.
            builder.InsertCell();
            builder.Write("Outer Cell 2");

            // Finish the outer table.
            builder.EndRow();
            builder.EndTable();

            // Create an inner table (2 rows, 2 columns) with sample text.
            Table innerTable = CreateTable(doc, 2, 2, "Inner");

            // Insert the inner table into the first cell of the outer table.
            outerTable.FirstRow.FirstCell.AppendChild(innerTable);

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "NestedTable.docx");
            doc.Save(outputPath);
        }

        // Helper method that creates a table with the specified dimensions and text.
        private static Table CreateTable(Document doc, int rowCount, int columnCount, string cellText)
        {
            Table table = new Table(doc);

            for (int r = 1; r <= rowCount; r++)
            {
                Row row = new Row(doc);
                table.AppendChild(row);

                for (int c = 1; c <= columnCount; c++)
                {
                    Cell cell = new Cell(doc);
                    // Ensure the cell contains at least one paragraph.
                    cell.AppendChild(new Paragraph(doc));
                    cell.FirstParagraph.AppendChild(new Run(doc, cellText));
                    row.AppendChild(cell);
                }
            }

            // Optional: set title/description for completeness.
            table.Title = "Inner Table";
            table.Description = "A nested table inside an outer table cell.";

            return table;
        }
    }
}
