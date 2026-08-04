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

        // First cell – set vertical alignment to middle (center).
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.Write("First cell with centered text.");

        // Second cell – also set vertical alignment to middle.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.Write("Second cell with centered text.");

        // End the first row.
        builder.EndRow();

        // Add a second row to demonstrate that previous rows are not affected.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1 (default alignment).");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2 (default alignment).");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Verify that the vertical alignment was applied.
        if (table.Rows[0].Cells[0].CellFormat.VerticalAlignment != CellVerticalAlignment.Center ||
            table.Rows[0].Cells[1].CellFormat.VerticalAlignment != CellVerticalAlignment.Center)
        {
            throw new InvalidOperationException("Vertical alignment was not set correctly.");
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TableWithVerticalAlignment.docx");
        doc.Save(outputPath);
    }
}
