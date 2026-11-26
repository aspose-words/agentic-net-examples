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

        // Build a simple table with two rows and two columns.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Short");
        builder.InsertCell();
        builder.Writeln("A much longer piece of text that should cause the column to expand.");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Another");
        builder.InsertCell();
        builder.Writeln("Content");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Enable automatic column width adjustment to fit the cell contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to disk.
        string outputPath = "TableAutoFit.docx";
        doc.Save(outputPath);

        // Verify that the file was saved successfully.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
