using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct a simple 2x2 table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Prevent the first row from breaking across pages.
        Row firstRow = table.FirstRow;
        firstRow.RowFormat.AllowBreakAcrossPages = false;

        // Save the document to the local file system.
        string outputPath = "TableAllowBreakAcrossPages.docx";
        doc.Save(outputPath);
    }
}
