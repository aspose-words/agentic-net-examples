using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableVerticalAlignmentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            builder.StartTable();

            // Define number of rows and columns.
            int rows = 3;
            int columns = 3;

            // Build the table and set vertical alignment for each cell.
            for (int row = 1; row <= rows; row++)
            {
                for (int col = 1; col <= columns; col++)
                {
                    // Insert a new cell.
                    builder.InsertCell();

                    // Add some text to the cell.
                    builder.Writeln($"Cell {row},{col}");

                    // Retrieve the current cell and set its vertical alignment to Bottom.
                    Cell currentCell = (Cell)builder.CurrentParagraph.ParentNode;
                    currentCell.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;
                }

                // End the current row.
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Save the document to a file.
            string outputPath = "VerticalAlignmentTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
