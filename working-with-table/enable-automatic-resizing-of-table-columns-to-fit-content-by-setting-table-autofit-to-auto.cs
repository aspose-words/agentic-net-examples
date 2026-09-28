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

        // Build a simple table with two rows and two columns.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Short");
        builder.InsertCell();
        builder.Writeln("This is a longer piece of text that should cause the column to expand.");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Another");
        builder.InsertCell();
        builder.Writeln("More content");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Enable automatic column width fitting to the cell contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to disk.
        string outputPath = "TableAutoFit.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!System.IO.File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
