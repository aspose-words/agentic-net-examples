using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableSplitExample
{
    public class Program
    {
        public static void Main()
        {
            // Define output folder and ensure it exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -------------------------------------------------
            // 1. Build a table with a horizontally merged cell.
            // -------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();

            // First cell – start of a merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged Cell");

            // Second cell – merged with the previous one.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // End the first row.
            builder.EndRow();

            // Add a second row with normal cells.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document that contains the merged cell.
            string mergedPath = Path.Combine(outputDir, "MergedCell.docx");
            doc.Save(mergedPath);

            // -------------------------------------------------
            // 2. Load the document and split the merged cell.
            // -------------------------------------------------
            Document loadedDoc = new Document(mergedPath);
            Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];

            // The merged cell is the first cell of the first row.
            Row firstRow = loadedTable.Rows[0];
            Cell mergedCell = firstRow.Cells[0];

            // Reset the merge flag on the original cell.
            mergedCell.CellFormat.HorizontalMerge = CellMerge.None;

            // Insert a new cell after the original one to replace the split part.
            Cell newCell = new Cell(loadedDoc);
            // Ensure the new cell has at least one paragraph (required for a valid cell).
            newCell.EnsureMinimum();

            // Insert the new cell into the row.
            firstRow.InsertAfter(newCell, mergedCell);

            // Save the document after splitting.
            string splitPath = Path.Combine(outputDir, "SplitCell.docx");
            loadedDoc.Save(splitPath);

            // Simple verification: output the cell count of the first row.
            int cellCount = firstRow.Cells.Count;
            Console.WriteLine($"First row now has {cellCount} cells after split.");
        }
    }
}
