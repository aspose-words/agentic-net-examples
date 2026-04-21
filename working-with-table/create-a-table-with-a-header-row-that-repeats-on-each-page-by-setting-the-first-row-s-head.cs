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

        // ----- Header row (repeated on each page) -----
        // Set HeadingFormat to true so this row repeats on every page the table spans.
        builder.RowFormat.HeadingFormat = true;

        // First header cell.
        builder.InsertCell();
        builder.Write("Header Column 1");

        // Second header cell.
        builder.InsertCell();
        builder.Write("Header Column 2");

        // Finish the header row.
        builder.EndRow();

        // ----- Data rows (regular rows) -----
        // Turn off heading format for subsequent rows.
        builder.RowFormat.HeadingFormat = false;

        // Add enough rows to make the table span multiple pages.
        for (int i = 1; i <= 60; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i}, Column 1");

            builder.InsertCell();
            builder.Write($"Row {i}, Column 2");

            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithRepeatingHeader.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
