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

        // Start a table and set explicit column widths.
        builder.StartTable();

        // First column width.
        builder.CellFormat.Width = 100; // points
        builder.InsertCell();
        builder.Writeln("First column");

        // Second column width.
        builder.CellFormat.Width = 150; // points
        builder.InsertCell();
        builder.Writeln("Second column");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Disable automatic column resizing while preserving the set widths.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Save the document.
        string outputPath = "TableAutoFit.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
