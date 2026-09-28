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

        // Build a simple table with 5 rows and 3 columns.
        builder.StartTable();

        for (int row = 1; row <= 5; row++)
        {
            for (int col = 1; col <= 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {row}, Cell {col}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Prevent rows from breaking across pages, which keeps the whole table on a single page.
        foreach (Row r in table.Rows)
        {
            r.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document.
        string outputPath = "TableKeepTogether.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException($"Failed to save the document to '{outputPath}'.");
    }
}
