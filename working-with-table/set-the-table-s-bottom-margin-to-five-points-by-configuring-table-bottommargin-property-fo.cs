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

        // Build a simple 2‑row, 1‑column table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set the table's bottom padding to five points (Aspose.Words uses padding, not margin).
        table.BottomPadding = 5.0;

        // Save the document to disk.
        string outputPath = "TableBottomMargin.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
