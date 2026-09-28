using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table with varying content length.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Short");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("A much longer text that should cause the column to expand");

        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Another short");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("More long content to test auto‑fit behavior");

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Enable automatic column resizing to fit the cell contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        string outputPath = "TableAutoFit.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
