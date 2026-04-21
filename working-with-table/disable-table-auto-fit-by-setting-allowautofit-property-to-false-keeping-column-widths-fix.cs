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
        Table table = builder.StartTable();

        // First row, first cell with a fixed width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("First column with fixed width.");

        // First row, second cell with a fixed width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Write("Second column with fixed width.");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Disable automatic resizing (auto‑fit) of the table.
        table.AllowAutoFit = false;

        // Save the document to the local file system.
        string outputPath = "TableAllowAutoFit.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to save the document.");

        Console.WriteLine($"Document saved successfully to {Path.GetFullPath(outputPath)}");
    }
}
