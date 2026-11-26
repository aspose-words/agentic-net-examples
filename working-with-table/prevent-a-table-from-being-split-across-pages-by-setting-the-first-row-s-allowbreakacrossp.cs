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

        // Build a table with enough rows to potentially span multiple pages.
        builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Additional rows with sample content.
        for (int i = 0; i < 30; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i + 1} Column 1");
            builder.InsertCell();
            builder.Writeln($"Row {i + 1} Column 2");
            builder.EndRow();
        }

        builder.EndTable();

        // Retrieve the table and set the first row to not break across pages.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Row firstRow = table.Rows[0];
        firstRow.RowFormat.AllowBreakAcrossPages = false;

        // Save the document.
        string outputPath = "TableNoBreak.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");
    }
}
