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

        // Header row.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Writeln("Data 1A");
        builder.InsertCell();
        builder.Writeln("Data 1B");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Writeln("Data 2A");
        builder.InsertCell();
        builder.Writeln("Data 2B");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Apply a built‑in style and enable the FirstRow option to format the header row.
        table.StyleIdentifier = StyleIdentifier.TableGrid;
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Save the document.
        string outputPath = "TableStyleFirstRow.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
