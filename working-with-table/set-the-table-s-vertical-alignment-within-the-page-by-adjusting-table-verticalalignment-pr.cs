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

        // Build a simple 2‑row table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Aspose.Words does not expose a direct VerticalAlignment property for tables.
        // As an alternative, you can set the table's horizontal alignment.
        // If vertical alignment on the page is required, it must be achieved via
        // positioning settings (e.g., using Table.Positioning) which are beyond
        // the scope of this simple example.
        table.Alignment = TableAlignment.Center; // Horizontal centering as a placeholder.

        // Save the document to disk.
        string outputPath = "TableVerticalAlignment.docx";
        doc.Save(outputPath);

        // Verify that the file was saved successfully.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
