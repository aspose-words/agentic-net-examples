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

        // -------------------------------------------------
        // Build a table with a horizontally merged cell.
        // -------------------------------------------------
        Table table = builder.StartTable();

        // First cell – start of the merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged Cell");

        // Second cell – merged with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Third cell – also merged with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first row.
        builder.EndRow();

        // Add a normal second row with three separate cells.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.InsertCell();
        builder.Write("Row 2, Cell 3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document that contains the merged cell.
        string mergedPath = Path.Combine(Environment.CurrentDirectory, "MergedTable.docx");
        doc.Save(mergedPath);

        // -------------------------------------------------
        // Split the previously merged cell back into separate cells.
        // -------------------------------------------------
        // The merged region consists of three cells in the first row.
        // To split them, clear the merge flags on each cell.
        Row firstRow = table.Rows[0];
        foreach (Cell cell in firstRow.Cells)
        {
            cell.CellFormat.HorizontalMerge = CellMerge.None;
            cell.CellFormat.VerticalMerge = CellMerge.None;
        }

        // Save the document after splitting.
        string splitPath = Path.Combine(Environment.CurrentDirectory, "SplitMergedCell.docx");
        doc.Save(splitPath);

        // Simple validation to ensure the files were created.
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged table document was not saved.", mergedPath);
        if (!File.Exists(splitPath))
            throw new FileNotFoundException("Split table document was not saved.", splitPath);

        // Indicate successful completion.
        Console.WriteLine("Documents created successfully:");
        Console.WriteLine($"- Merged table: {mergedPath}");
        Console.WriteLine($"- Split table: {splitPath}");
    }
}
