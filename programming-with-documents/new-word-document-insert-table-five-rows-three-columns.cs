using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The builder returns the created Table object.
        Table table = builder.StartTable();

        // Insert 5 rows.
        for (int row = 1; row <= 5; row++)
        {
            // Insert 3 columns (cells) for each row.
            for (int col = 1; col <= 3; col++)
            {
                // Insert a new cell into the current row.
                builder.InsertCell();

                // Write placeholder text into the cell.
                builder.Write($"Row {row}, Col {col}");
            }

            // End the current row and move to the next one.
            builder.EndRow();
        }

        // Close the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("TableFiveRowsThreeColumns.docx");
    }
}
