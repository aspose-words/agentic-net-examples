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
        builder.Write("First cell (centered vertically).");

        // Second cell – also set vertical alignment to middle.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.Write("Second cell (centered vertically).");

        // End the first row.
        builder.EndRow();

        // Add a second row with default alignment for contrast.
        builder.InsertCell();
        builder.Write("Third cell (default alignment).");
        builder.InsertCell();
        builder.Write("Fourth cell (default alignment).");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "VerticalAlignmentTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // The program ends automatically; no user interaction required.
    }
}
