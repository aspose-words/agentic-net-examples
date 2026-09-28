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
        builder.StartTable();

        // Insert first row (header) with three cells.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.EndRow();

        // Insert several more rows to increase the table height.
        for (int i = 1; i <= 20; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i} Col 1");
            builder.InsertCell();
            builder.Writeln($"Row {i} Col 2");
            builder.InsertCell();
            builder.Writeln($"Row {i} Col 3");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the first row and prevent it from breaking across pages.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Row firstRow = table.Rows[0];
        // The AllowBreakAcrossPages property belongs to RowFormat, not Row itself.
        firstRow.RowFormat.AllowBreakAcrossPages = false;

        // Save the document.
        string outputPath = "TableNoBreak.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
