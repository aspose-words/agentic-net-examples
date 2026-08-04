using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table that will likely span multiple pages.
        builder.StartTable();

        // Optional header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Add many rows to the table.
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

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Prevent each row from breaking across pages.
        foreach (Row row in table.Rows)
        {
            // Setting AllowBreakAcrossPages to false keeps the row together.
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to the local file system.
        doc.Save("PreventRowBreak.docx");
    }
}
