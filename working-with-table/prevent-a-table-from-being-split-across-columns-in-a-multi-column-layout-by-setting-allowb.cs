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

        // Set the page to have two columns (multi‑column layout).
        builder.PageSetup.TextColumns.SetCount(2);

        // Add some introductory text.
        builder.Writeln("Text before the table.");

        // Start building a table.
        Table table = builder.StartTable();

        // Create a simple 3‑row, 2‑column table.
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                builder.InsertCell();
                builder.Write($"Row {row + 1}, Cell {col + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Prevent each row from breaking across columns (or pages) in the multi‑column layout.
        foreach (Row r in table.Rows)
        {
            r.RowFormat.AllowBreakAcrossPages = false;
        }

        // Add some text after the table.
        builder.Writeln("Text after the table.");

        // Save the document to a file.
        doc.Save("Table_NoBreakAcrossColumns.docx");
    }
}
