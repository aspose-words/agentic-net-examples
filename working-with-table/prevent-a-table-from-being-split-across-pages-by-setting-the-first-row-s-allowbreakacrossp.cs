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

        // Start a table and add a header row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Add enough rows to make the table span multiple pages.
        for (int i = 0; i < 50; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i + 1} Column 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1} Column 2");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Prevent the first row from breaking across pages.
        Row firstRow = table.FirstRow;
        firstRow.RowFormat.AllowBreakAcrossPages = false;

        // Save the document.
        string outputPath = "Table_NoBreakAcrossPages.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved correctly.");
    }
}
