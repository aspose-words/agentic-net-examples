using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a table.
        Table table = builder.StartTable();

        int rows = 3;
        int columns = 2;

        // Build the table rows and cells.
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                // Set vertical alignment to Bottom for each new cell.
                builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;

                // Insert the cell and add some sample text.
                builder.InsertCell();
                builder.Write($"Row {r + 1}, Col {c + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("CellVerticalAlignmentBottom.docx");
    }
}
