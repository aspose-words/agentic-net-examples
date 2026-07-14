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
        builder.StartTable();

        // Insert 5 rows and 3 columns.
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                // Insert a cell and add some text.
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to a file in the current directory.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "Table.docx");
        doc.Save(outputPath);
    }
}
