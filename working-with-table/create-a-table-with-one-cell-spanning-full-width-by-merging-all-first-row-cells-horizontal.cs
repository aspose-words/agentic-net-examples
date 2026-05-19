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
        Table table = builder.StartTable();

        // ---- First row: one cell merged across the full width ----
        // Insert the first cell and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("This cell spans the entire first row.");

        // Insert additional cells that will be merged with the first one.
        // The number of extra cells determines the original column count.
        int extraCells = 3; // Adjust as needed for the desired width.
        for (int i = 0; i < extraCells; i++)
        {
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        }

        // End the first row.
        builder.EndRow();

        // ---- Second row: normal cells (optional, just to show the table continues) ----
        builder.InsertCell();
        builder.Write("Second row, cell 1");
        builder.InsertCell();
        builder.Write("Second row, cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
