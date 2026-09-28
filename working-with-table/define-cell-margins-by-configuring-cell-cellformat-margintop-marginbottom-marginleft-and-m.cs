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

        // Start a new table.
        builder.StartTable();

        // Insert first cell.
        builder.InsertCell();
        builder.Writeln("Cell with custom margins");

        // Retrieve the cell that was just created.
        Cell cellWithMargins = (Cell)builder.CurrentParagraph.ParentNode;

        // Set custom margins (in points). 1 point = 1/72 inch.
        // In Aspose.Words the cell padding properties are used for margins.
        cellWithMargins.CellFormat.TopPadding = 10;      // 10 points top margin
        cellWithMargins.CellFormat.BottomPadding = 10;   // 10 points bottom margin
        cellWithMargins.CellFormat.LeftPadding = 15;     // 15 points left margin
        cellWithMargins.CellFormat.RightPadding = 15;    // 15 points right margin

        // Insert second cell with default margins for comparison.
        builder.InsertCell();
        builder.Writeln("Cell with default margins");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "CellMargins.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
