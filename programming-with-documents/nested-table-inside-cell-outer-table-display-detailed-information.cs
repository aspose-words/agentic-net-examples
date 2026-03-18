using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsExamples
{
    public class NestedTableDemo
    {
        public static void CreateNestedTable()
        {
            // Initialize a new blank document.
            Document doc = new Document();

            // Build the outer table (e.g., 3 rows x 4 columns) with placeholder text.
            Table outerTable = CreateTable(doc, 3, 4, "Outer Table");
            // Append the outer table to the document body.
            doc.FirstSection.Body.AppendChild(outerTable);

            // Build the inner (nested) table (e.g., 2 rows x 2 columns) with its own placeholder text.
            Table innerTable = CreateTable(doc, 2, 2, "Inner Table");
            // Insert the inner table into the first cell of the outer table.
            outerTable.FirstRow.FirstCell.AppendChild(innerTable);

            // Save the document to disk.
            doc.Save("NestedTable.docx");
        }

        /// <summary>
        /// Helper method that creates a table with the specified number of rows and columns.
        /// Each cell will contain a single paragraph with the supplied text.
        /// </summary>
        private static Table CreateTable(Document doc, int rowCount, int columnCount, string cellText)
        {
            Table table = new Table(doc);

            // Populate the table.
            for (int r = 1; r <= rowCount; r++)
            {
                Row row = new Row(doc);
                table.AppendChild(row);

                for (int c = 1; c <= columnCount; c++)
                {
                    Cell cell = new Cell(doc);
                    // Ensure the cell has a paragraph to host the text.
                    cell.AppendChild(new Paragraph(doc));
                    // Add the text run.
                    cell.FirstParagraph.AppendChild(new Run(doc, cellText));
                    row.AppendChild(cell);
                }
            }

            // Optional: set title/description for better accessibility.
            table.Title = "Aspose table title";
            table.Description = "Aspose table description";

            return table;
        }

        // Entry point for testing.
        public static void Main()
        {
            CreateNestedTable();
            Console.WriteLine("Nested table document created successfully.");
        }
    }
}
