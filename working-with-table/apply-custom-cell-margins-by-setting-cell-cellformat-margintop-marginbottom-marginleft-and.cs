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

        // Insert the first cell.
        Cell cell = builder.InsertCell();

        // Apply custom margins (implemented via cell padding properties).
        // These values are in points.
        cell.CellFormat.TopPadding = 10;    // MarginTop
        cell.CellFormat.BottomPadding = 15; // MarginBottom
        cell.CellFormat.LeftPadding = 20;   // MarginLeft
        cell.CellFormat.RightPadding = 25;  // MarginRight

        // Add some text to the cell.
        builder.Write("Cell with custom margins.");

        // Insert a second cell with default margins for comparison.
        builder.InsertCell();
        builder.Write("Cell with default margins.");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "CustomCellMargins.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");

        // Optionally, you could open the file automatically (commented out to avoid side effects).
        // System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
    }
}
