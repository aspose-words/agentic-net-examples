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

        // Mark the first row as a heading row that repeats on each page.
        builder.RowFormat.HeadingFormat = true;

        // First header cell.
        builder.InsertCell();
        builder.Write("Header Column 1");

        // Second header cell.
        builder.InsertCell();
        builder.Write("Header Column 2");
        builder.EndRow();

        // Subsequent rows should not repeat as headings.
        builder.RowFormat.HeadingFormat = false;

        // Add enough rows to make the table span multiple pages.
        for (int i = 1; i <= 50; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i}, Column 1");
            builder.InsertCell();
            builder.Write($"Row {i}, Column 2");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Ensure the output directory exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithRepeatingHeader.docx");
        doc.Save(outputPath);
    }
}
