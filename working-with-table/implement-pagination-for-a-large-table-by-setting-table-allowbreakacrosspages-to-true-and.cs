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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        builder.StartTable();

        // Define the number of rows and columns for the large table.
        int rowCount = 50;
        int columnCount = 5;

        // Populate the table with sample data.
        for (int i = 0; i < rowCount; i++)
        {
            for (int j = 0; j < columnCount; j++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {i + 1}, Cell {j + 1}");
            }
            // End the current row.
            builder.EndRow();
        }

        // Finish the table construction.
        builder.EndTable();

        // Retrieve the created table from the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Enable pagination for each row so the table can break across pages.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = true;

            // Adjust the height of each row to a fixed value.
            row.RowFormat.Height = 20; // Height in points.
            row.RowFormat.HeightRule = HeightRule.Exactly;
        }

        // Save the document to a file.
        string outputPath = "TablePagination.docx";
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
