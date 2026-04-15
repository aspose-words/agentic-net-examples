using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Build a table with 5 rows and 3 columns.
        for (int row = 1; row <= 5; row++)
        {
            for (int col = 1; col <= 3; col++)
            {
                // Insert a new cell and write sample text.
                builder.InsertCell();
                builder.Write($"Row {row}, Column {col}");
            }

            // End the current row before starting the next one.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to a file named "Table.docx" in the current directory.
        doc.Save("Table.docx");
    }
}
