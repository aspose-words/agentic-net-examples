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
        Table table = builder.StartTable();

        // Populate the table with a few rows and cells.
        for (int i = 0; i < 5; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Cell 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Cell 2");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Configure each row so it cannot break across pages.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Table_NoBreakAcrossPages.docx");
        doc.Save(outputPath);
    }
}
