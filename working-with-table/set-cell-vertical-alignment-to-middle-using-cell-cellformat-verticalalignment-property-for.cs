using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Path where the resulting document will be saved.
        string outputPath = "VerticalAlignmentTable.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // First row, first cell – set vertical alignment to middle (center).
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.Write("Cell 1\r\nLine 2\r\nLine 3");

        // First row, second cell – also middle-aligned.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.Write("Cell 2");

        // End the first row.
        builder.EndRow();

        // Second row – increase row height to demonstrate vertical centering.
        builder.RowFormat.Height = 100;
        builder.RowFormat.HeightRule = HeightRule.Exactly;

        // Second row, first cell.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.Write("Second row, cell 1");

        // Second row, second cell.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.Write("Second row, cell 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
