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

        // Build a simple 3‑row, 2‑column table.
        builder.StartTable();

        for (int i = 1; i <= 3; i++)
        {
            // First cell of the row.
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 1");

            // Second cell of the row.
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 2");

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = doc.FirstSection.Body.Tables[0];

        // Disable breaking of each row across pages.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to disk.
        string outputPath = "Table_NoBreakAcrossPages.docx";
        doc.Save(outputPath);
    }
}
