using System;
using System.IO;
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

        // Begin the table.
        builder.StartTable();

        int rows = 1000;   // Large number of rows.
        int columns = 5;   // Number of columns.

        // Populate the table with cells containing simple text.
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < columns; j++)
            {
                builder.InsertCell();
                builder.Writeln($"R{i + 1}C{j + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Force a layout pass after all modifications.
        doc.UpdatePageLayout();

        // Save the document.
        string outputPath = "OptimizedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
