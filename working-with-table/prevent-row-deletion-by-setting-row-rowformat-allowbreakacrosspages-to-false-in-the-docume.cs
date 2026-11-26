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

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Prevent rows from breaking across pages.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to disk.
        string outputPath = "PreventRowDeletion.docx";
        doc.Save(outputPath);

        // Simple validation that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
