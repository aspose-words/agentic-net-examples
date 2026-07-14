using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Build a 2x2 table where each cell's text is aligned to the bottom.
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                // Apply bottom vertical alignment to the next cell.
                builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;

                // Insert the cell and write some text.
                builder.InsertCell();
                builder.Write($"Row {row + 1}, Cell {col + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("CellVerticalAlignmentBottom.docx");
    }
}
