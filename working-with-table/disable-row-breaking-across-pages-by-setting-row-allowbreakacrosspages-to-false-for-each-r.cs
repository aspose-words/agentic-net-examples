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

        // Build a simple table with several rows.
        Table table = builder.StartTable();

        for (int i = 1; i <= 5; i++)
        {
            // First cell of the row.
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 1. This is some sample text to make the row relatively long.");

            // Second cell of the row.
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 2. Additional content to increase the row height.");

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Disable breaking rows across pages for every row in the table.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to disk.
        doc.Save("TableAllowBreakAcrossPages.docx");
    }
}
