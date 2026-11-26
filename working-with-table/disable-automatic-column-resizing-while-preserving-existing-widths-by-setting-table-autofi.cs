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

        // First column – set a fixed width.
        builder.CellFormat.Width = 100; // Width in points.
        builder.InsertCell();
        builder.Writeln("Column 1");

        // Second column – set a fixed width.
        builder.CellFormat.Width = 150;
        builder.InsertCell();
        builder.Writeln("Column 2");

        // End the first row.
        builder.EndRow();

        // Add a second row with the same column widths.
        builder.InsertCell();
        builder.Writeln("Row 2, Col 1");
        builder.InsertCell();
        builder.Writeln("Row 2, Col 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Disable automatic column resizing while preserving the existing widths.
        // Use the AutoFit method with FixedColumnWidths behavior.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Save the document.
        string outputPath = "TableAutoFit.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
