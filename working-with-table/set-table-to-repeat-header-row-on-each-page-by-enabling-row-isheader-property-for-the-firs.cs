using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "TableWithRepeatingHeader.docx");

        // Create a new blank document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Mark the first row as a heading row that repeats on each page.
        builder.RowFormat.HeadingFormat = true;

        // Build the header row.
        builder.InsertCell();
        builder.Write("Header Column 1");
        builder.InsertCell();
        builder.Write("Header Column 2");
        builder.EndRow();

        // Subsequent rows should not repeat.
        builder.RowFormat.HeadingFormat = false;

        // Add enough rows to make the table span multiple pages.
        for (int i = 1; i <= 50; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} Column 1");
            builder.InsertCell();
            builder.Write($"Row {i} Column 2");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved.");

        // Verify that the first row is set to repeat as a header.
        Table savedTable = doc.FirstSection.Body.Tables[0];
        if (!savedTable.FirstRow.RowFormat.HeadingFormat)
            throw new Exception("The header row was not configured correctly.");
    }
}
