using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableColumnWidthExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document and a DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Sample data for the table.
            string[,] data = new string[,]
            {
                { "ID", "Name", "Description" },
                { "1", "Alice", "Software engineer with 5 years experience" },
                { "2", "Bob", "Data analyst" },
                { "3", "Charlie", "Project manager overseeing multiple teams" }
            };

            // Build the table.
            builder.StartTable();
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    builder.InsertCell();
                    builder.Writeln(data[r, c]);
                }
                builder.EndRow();
            }
            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

            // Determine the maximum text length for each column.
            int[] maxLengths = new int[cols];
            foreach (Row row in table.Rows)
            {
                for (int c = 0; c < cols; c++)
                {
                    Cell cell = row.Cells[c];
                    // Get the plain text of the cell (trim trailing cell marker).
                    string cellText = cell.GetText().TrimEnd('\a');
                    if (cellText.Length > maxLengths[c])
                        maxLengths[c] = cellText.Length;
                }
            }

            // Approximate width: assume 7 points per character (adjust as needed).
            const double pointsPerChar = 7.0;

            // Apply the calculated width to each cell in the column.
            for (int c = 0; c < cols; c++)
            {
                double columnWidth = maxLengths[c] * pointsPerChar;
                foreach (Row row in table.Rows)
                {
                    Cell cell = row.Cells[c];
                    cell.CellFormat.Width = columnWidth;
                }
            }

            // Save the document.
            string outputPath = "TableColumnWidths.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
