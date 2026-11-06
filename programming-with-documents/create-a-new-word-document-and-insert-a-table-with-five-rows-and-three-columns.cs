using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Insert 5 rows and 3 columns.
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                // Insert a cell with sample text.
                builder.InsertCell();
                builder.Writeln($"Row {row + 1}, Col {col + 1}");
            }

            // End the current row and start a new one.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "TableDocument.docx";
        doc.Save(outputPath);
    }
}
