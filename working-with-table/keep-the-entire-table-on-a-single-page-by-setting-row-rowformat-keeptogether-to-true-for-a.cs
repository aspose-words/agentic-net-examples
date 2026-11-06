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

        for (int rowIndex = 0; rowIndex < 5; rowIndex++)
        {
            for (int colIndex = 0; colIndex < 3; colIndex++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {rowIndex + 1}, Cell {colIndex + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Prevent rows from breaking across pages, which keeps the whole table on one page.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document.
        string outputPath = "TableKeepTogether.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to save the document to '{outputPath}'.");
        }
    }
}
