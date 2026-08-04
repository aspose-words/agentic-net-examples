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

        // Start building a table.
        Table table = builder.StartTable();

        // Define the size of the large table.
        int rowCount = 50;   // Number of rows.
        int colCount = 5;    // Number of columns.

        // Populate the table.
        for (int i = 0; i < rowCount; i++)
        {
            // Insert cells for the current row.
            for (int j = 0; j < colCount; j++)
            {
                builder.InsertCell();
                builder.Write($"Row {i + 1}, Col {j + 1}");
            }

            // End the current row and obtain the Row object.
            Row row = builder.EndRow();

            // Allow the row to break across pages.
            row.RowFormat.AllowBreakAcrossPages = true;

            // Set a minimum height for the row to ensure consistent layout.
            row.RowFormat.Height = 20;               // Height in points.
            row.RowFormat.HeightRule = HeightRule.AtLeast;
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "LargeTablePagination.docx";
        doc.Save(outputPath);
    }
}
