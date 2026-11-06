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

        // Add a paragraph before the table.
        builder.Writeln("Paragraph before the table.");

        // Build a simple 1x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Add a paragraph after the table.
        builder.Writeln("Paragraph after the table.");

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Set the table to wrap text around it.
        // The WrapAround property is not available in this version; use TextWrapping instead.
        table.TextWrapping = TextWrapping.Around;

        // Save the document.
        string outputPath = "TableWrapAround.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
