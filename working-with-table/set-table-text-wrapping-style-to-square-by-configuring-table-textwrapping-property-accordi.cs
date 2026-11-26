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

        // Build a simple 2x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set the table's text wrapping style.
        // The TextWrapping enum in the current Aspose.Words version provides
        // None, Around, and TopAndBottom. The "Square" option is not available,
        // so we use the closest equivalent (Around) which wraps text around the table.
        table.TextWrapping = TextWrapping.Around;

        // Save the document to disk.
        string outputPath = "TableTextWrapping.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!System.IO.File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
