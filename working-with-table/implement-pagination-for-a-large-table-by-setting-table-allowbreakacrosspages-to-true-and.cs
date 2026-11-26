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

        // Create a large table with many rows to demonstrate pagination.
        const int totalRows = 50;
        const int columns = 3;

        for (int rowIdx = 1; rowIdx <= totalRows; rowIdx++)
        {
            for (int colIdx = 1; colIdx <= columns; colIdx++)
            {
                // Insert cell text.
                builder.InsertCell();
                builder.Writeln($"Row {rowIdx}, Column {colIdx}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Allow each row to break across pages.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = true;

            // Adjust each row's height to a fixed value.
            row.RowFormat.Height = 20.0;               // Height in points.
            row.RowFormat.HeightRule = HeightRule.Exactly;
        }

        // Save the document.
        string outputPath = "TablePagination.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
