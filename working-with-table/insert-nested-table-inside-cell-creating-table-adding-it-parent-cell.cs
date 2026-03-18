using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace NestedTableExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Build the outer table (3 rows x 4 columns) and add it to the document body.
            Table outerTable = CreateTable(doc, 3, 4, "Outer Table");
            doc.FirstSection.Body.AppendChild(outerTable);

            // Build the inner table (2 rows x 2 columns).
            Table innerTable = CreateTable(doc, 2, 2, "Inner Table");

            // Insert the inner table into the first cell of the outer table.
            outerTable.FirstRow.FirstCell.AppendChild(innerTable);

            // Save the document.
            doc.Save("NestedTable.docx");
        }

        /// <summary>
        /// Creates a table with the specified number of rows and columns.
        /// Each cell contains a single paragraph with the supplied text.
        /// </summary>
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
                    // Ensure the cell has at least one paragraph.
                    cell.EnsureMinimum();

                    // Add the desired text to the cell.
                    cell.FirstParagraph.AppendChild(new Run(doc, cellText));

                    row.AppendChild(cell);
                }
            }

            // Optional: set title/description for ISO/IEC 29500 compliance.
            table.Title = "Aspose table title";
            table.Description = "Aspose table description";

            return table;
        }
    }
}
