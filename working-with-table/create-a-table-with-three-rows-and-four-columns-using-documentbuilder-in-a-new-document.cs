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
        builder.StartTable();

        // Build 3 rows × 4 columns.
        for (int row = 1; row <= 3; row++)
        {
            for (int col = 1; col <= 4; col++)
            {
                // Insert a new cell and write some text into it.
                builder.InsertCell();
                builder.Write($"R{row}C{col}");
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
