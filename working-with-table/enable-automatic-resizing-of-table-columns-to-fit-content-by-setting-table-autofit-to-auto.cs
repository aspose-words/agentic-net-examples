using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with two rows and three columns.
        builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Writeln("Short");
        builder.InsertCell();
        builder.Writeln("A bit longer text");
        builder.InsertCell();
        builder.Writeln("This is a considerably longer piece of text to demonstrate auto‑fit.");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table (the last table in the document).
        Table table = doc.GetChildNodes(NodeType.Table, true)
                         .Cast<Table>()
                         .LastOrDefault();

        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Enable automatic column resizing to fit the content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to disk.
        string outputPath = "AutoFitTable.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
