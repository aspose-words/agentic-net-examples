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
                // Insert a new cell.
                builder.InsertCell();

                // Add some sample text to the cell.
                builder.Writeln($"Row {row + 1}, Col {col + 1}");

                // End the row after the last column.
                if (col == 2)
                {
                    builder.EndRow();
                }
            }
        }

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "TableDocument.docx";
        doc.Save(outputPath);
    }
}
