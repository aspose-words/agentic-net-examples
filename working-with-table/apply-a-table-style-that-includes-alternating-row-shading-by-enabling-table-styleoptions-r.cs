using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with a header row and several data rows.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Data rows.
        for (int i = 1; i <= 5; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i} Col 1");
            builder.InsertCell();
            builder.Writeln($"Row {i} Col 2");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a built‑in style and enable alternating row shading.
        table.StyleIdentifier = StyleIdentifier.TableGrid;
        table.StyleOptions = TableStyleOptions.RowBands;

        // Save the document.
        string outputPath = "TableRowBanding.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not created.");
        }
    }
}
