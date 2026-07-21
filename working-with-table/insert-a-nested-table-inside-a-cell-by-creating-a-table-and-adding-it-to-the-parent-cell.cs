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

            // Build the outer table using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder(doc);
            Table outerTable = builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.Write("Outer Cell 1");

            // First row, second cell.
            builder.InsertCell();
            builder.Write("Outer Cell 2");
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.Write("Outer Cell 3");

            // Second row, second cell.
            builder.InsertCell();
            builder.Write("Outer Cell 4");
            builder.EndRow();

            // Finish the outer table.
            builder.EndTable();

            // Create an inner table (2 rows x 2 columns) using a helper method.
            Table innerTable = CreateTable(doc, 2, 2, "Inner");

            // Insert the inner table into the first cell of the outer table.
            outerTable.FirstRow.FirstCell.AppendChild(innerTable);

            // Define the output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "NestedTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }

        /// <summary>
        /// Creates a table with the specified number of rows and columns.
        /// Each cell contains a paragraph with the provided text.
        /// </summary>
        private static Table CreateTable(Document doc, int rowCount, int columnCount, string cellText)
        {
            Table table = new Table(doc);

            for (int r = 0; r < rowCount; r++)
            {
                Row row = new Row(doc);
                table.AppendChild(row);

                for (int c = 0; c < columnCount; c++)
                {
                    Cell cell = new Cell(doc);
                    // Ensure the cell has at least one paragraph.
                    cell.AppendChild(new Paragraph(doc));
                    cell.FirstParagraph.AppendChild(new Run(doc, $"{cellText} ({r + 1},{c + 1})"));
                    row.AppendChild(cell);
                }
            }

            // Optional: set title/description for completeness.
            table.Title = "Inner Table";
            table.Description = "A nested table example";

            return table;
        }
    }
}
