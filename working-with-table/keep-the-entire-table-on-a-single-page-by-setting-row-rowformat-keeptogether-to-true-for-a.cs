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

        // Build a simple table with 5 rows and 2 columns.
        Table table = builder.StartTable();

        for (int rowIndex = 0; rowIndex < 5; rowIndex++)
        {
            // First cell of the row.
            builder.InsertCell();
            builder.Write($"Row {rowIndex + 1}, Cell 1");

            // Second cell of the row.
            builder.InsertCell();
            builder.Write($"Row {rowIndex + 1}, Cell 2");

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Ensure that each row is kept together on the same page.
        // The RowFormat does not expose a KeepTogether property, but setting
        // AllowBreakAcrossPages to false prevents the row from splitting across pages.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableKeepTogether.docx");
        doc.Save(outputPath);
    }
}
