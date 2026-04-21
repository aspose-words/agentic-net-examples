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
        builder.Write("Cell 1\nLine 2");

        // Second cell – also set vertical alignment to middle.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.Write("Cell 2");

        // End the first row.
        builder.EndRow();

        // Add a second row without changing vertical alignment (defaults to top) for comparison.
        builder.InsertCell();
        builder.Write("Normal top");
        builder.InsertCell();
        builder.Write("Normal top");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Validate that the first row cells have the expected vertical alignment.
        if (table.Rows[0].Cells[0].CellFormat.VerticalAlignment != CellVerticalAlignment.Center ||
            table.Rows[0].Cells[1].CellFormat.VerticalAlignment != CellVerticalAlignment.Center)
        {
            throw new InvalidOperationException("Vertical alignment was not applied correctly.");
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TableWithVerticalAlignment.docx");
        doc.Save(outputPath);
    }
}
