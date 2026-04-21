using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMerge
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // ----- First row: merge three cells horizontally -----
            // First cell – start of a merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged across three cells.");

            // Second cell – merge with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Third cell – also merge with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // End the first row.
            builder.EndRow();

            // ----- Second row: normal, unmerged cells (for comparison) -----
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Cell 1");

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Cell 2");

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Cell 3");

            // End the second row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCells.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not created.");

            // Load the saved document.
            Document loadedDoc = new Document(outputPath);
            Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];

            // Convert width‑based merges (if any) to merge flags.
            loadedTable.ConvertToHorizontallyMergedCells();

            // Validate that the first cell has the expected HorizontalMerge flag.
            Cell firstCell = loadedTable.Rows[0].Cells[0];
            if (firstCell.CellFormat.HorizontalMerge != CellMerge.First)
                throw new Exception("The first cell does not have the expected HorizontalMerge flag.");

            // Execution completed successfully.
        }
    }
}
