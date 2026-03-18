using Aspose.Words;
using Aspose.Words.Tables;
using System;

class Program
{
    static void Main()
    {
        // Create a new empty document and attach a DocumentBuilder to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the desired number of columns and rows for the table.
        int columnCount = 3;
        int rowCount = 4;

        // Start a new table. The method returns the Table node that was created.
        Table table = builder.StartTable();

        // Build the table row by row.
        for (int row = 0; row < rowCount; row++)
        {
            // Insert the required number of cells for each row.
            for (int col = 0; col < columnCount; col++)
            {
                builder.InsertCell();                                   // Insert a new cell.
                builder.Write($"Row {row + 1}, Col {col + 1}");        // Write sample text into the cell.
            }

            // End the current row before starting the next one.
            builder.EndRow();
        }

        // Finish the table. This moves the cursor out of the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("TableExample.docx");
    }
}
