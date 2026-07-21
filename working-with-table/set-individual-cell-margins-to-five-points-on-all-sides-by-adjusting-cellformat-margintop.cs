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

        // First row - first cell.
        builder.InsertCell();
        // Set cell margins (padding) to 5 points on all sides.
        builder.CellFormat.TopPadding = 5;
        builder.CellFormat.BottomPadding = 5;
        builder.CellFormat.LeftPadding = 5;
        builder.CellFormat.RightPadding = 5;
        builder.Write("Cell 1, Row 1");

        // First row - second cell.
        builder.InsertCell();
        builder.CellFormat.TopPadding = 5;
        builder.CellFormat.BottomPadding = 5;
        builder.CellFormat.LeftPadding = 5;
        builder.CellFormat.RightPadding = 5;
        builder.Write("Cell 2, Row 1");

        // End the first row.
        builder.EndRow();

        // Second row - first cell.
        builder.InsertCell();
        builder.CellFormat.TopPadding = 5;
        builder.CellFormat.BottomPadding = 5;
        builder.CellFormat.LeftPadding = 5;
        builder.CellFormat.RightPadding = 5;
        builder.Write("Cell 1, Row 2");

        // Second row - second cell.
        builder.InsertCell();
        builder.CellFormat.TopPadding = 5;
        builder.CellFormat.BottomPadding = 5;
        builder.CellFormat.LeftPadding = 5;
        builder.CellFormat.RightPadding = 5;
        builder.Write("Cell 2, Row 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellMargins.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");

        // Optionally, inform that the process completed (no interactive input required).
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
