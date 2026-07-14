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

        // Start a new table.
        Table table = builder.StartTable();

        // Build a large table (e.g., 100 rows) to demonstrate pagination.
        for (int i = 1; i <= 100; i++)
        {
            // First cell of the row.
            builder.InsertCell();
            builder.Writeln($"Row {i}, Cell 1");

            // Second cell of the row.
            builder.InsertCell();
            builder.Writeln($"Row {i}, Cell 2");

            // Adjust row formatting:
            // - Allow the row to break across pages.
            // - Set a minimum height so rows are tall enough to be noticeable.
            builder.RowFormat.AllowBreakAcrossPages = true;
            builder.RowFormat.Height = 20;               // Height in points.
            builder.RowFormat.HeightRule = HeightRule.AtLeast;

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LargeTablePagination.docx");
        doc.Save(outputPath);
    }
}
