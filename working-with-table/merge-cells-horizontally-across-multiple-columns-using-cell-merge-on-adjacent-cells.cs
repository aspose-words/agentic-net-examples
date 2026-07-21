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

        // ---------- First row: merge three cells horizontally ----------
        // Insert the first cell and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("This cell spans three columns.");

        // Insert the second cell and merge it with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Insert the third cell and merge it with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first row.
        builder.EndRow();

        // Reset merge settings for subsequent rows.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // ---------- Second row: normal (unmerged) cells ----------
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Define the output file path (in the same folder as the executable).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Optionally, inform that the process completed.
        Console.WriteLine("Document created successfully at: " + outputPath);
    }
}
