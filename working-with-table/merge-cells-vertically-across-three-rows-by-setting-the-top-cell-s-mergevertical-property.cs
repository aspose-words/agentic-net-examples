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

        // Start a table with three rows and two columns.
        Table table = builder.StartTable();

        // ----- Row 1 -----
        // First cell – start of a vertically merged range (covers three rows).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged vertically across three rows.");

        // Second cell – normal (no merging).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 1, Cell 2.");
        builder.EndRow();

        // ----- Row 2 -----
        // First cell – continues the vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Second cell – normal.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 2, Cell 2.");
        builder.EndRow();

        // ----- Row 3 -----
        // First cell – continues the vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Second cell – normal.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 3, Cell 2.");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Define output path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Optionally, inform that the process completed.
        Console.WriteLine("Document created successfully at: " + outputPath);
    }
}
