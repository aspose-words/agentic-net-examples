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

        // Build a simple 2x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("First cell");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Second cell");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new Exception("Table was not created.");

        // Set the table's horizontal alignment to the center of the page.
        // (The current Aspose.Words version does not expose a direct vertical alignment property for tables.)
        table.Alignment = TableAlignment.Center;

        // Save the document.
        string outputPath = "TableAlignment.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved.");

        // Reload the document and confirm the alignment.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = (Table)loadedDoc.GetChild(NodeType.Table, 0, true);
        if (loadedTable.Alignment != TableAlignment.Center)
            throw new Exception("Table alignment was not set correctly.");
    }
}
