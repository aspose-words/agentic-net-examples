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

        // Build a table with several rows and two cells per row.
        builder.StartTable();

        for (int rowIndex = 1; rowIndex <= 5; rowIndex++)
        {
            // First cell.
            builder.InsertCell();
            builder.Writeln($"Row {rowIndex}, Cell 1");

            // Second cell.
            builder.InsertCell();
            builder.Writeln($"Row {rowIndex}, Cell 2");

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Prevent each row from breaking across pages.
        foreach (Row row in table.Rows)
        {
            // Setting AllowBreakAcrossPages to false keeps the row together.
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to a file.
        string outputPath = "KeepTogetherTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
