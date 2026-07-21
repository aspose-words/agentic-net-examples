using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to simplify adding content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // Insert 5 rows and 3 columns.
        for (int row = 1; row <= 5; row++)
        {
            for (int col = 1; col <= 3; col++)
            {
                // Insert a new cell and write some text into it.
                builder.InsertCell();
                builder.Write($"Row {row}, Col {col}");
            }

            // End the current row before starting the next one.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to a file in the current directory.
        doc.Save("TableExample.docx");
    }
}
