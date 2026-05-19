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

        // Start a table.
        Table table = builder.StartTable();

        // Prevent rows from breaking across pages.
        builder.RowFormat.AllowBreakAcrossPages = false;

        // Add rows that will likely span multiple pages.
        for (int i = 1; i <= 30; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i}, Column 1");
            builder.InsertCell();
            builder.Writeln($"Row {i}, Column 2");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RowsKeepTogether.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
