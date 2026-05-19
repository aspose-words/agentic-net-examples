using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct the table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartTable();

        // Insert 5 rows and 3 columns.
        for (int row = 1; row <= 5; row++)
        {
            for (int col = 1; col <= 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row}C{col}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("TableExample.docx");
    }
}
