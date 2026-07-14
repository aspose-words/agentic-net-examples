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

        // Insert the first cell and mark it as the first cell in a horizontally merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged cell content");

        // Insert the second adjacent cell and merge it with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text is written to the merged cell.

        // End the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCells.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }

        // Optionally, inform that the operation completed successfully.
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
