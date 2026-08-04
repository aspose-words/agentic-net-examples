using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table with two columns.
            Table table = builder.StartTable();

            // ----- First Row -----
            // First column cell – this will become the first cell of the vertically merged range.
            builder.InsertCell();
            // Mark this cell as the first in a vertical merge.
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Merged vertically");

            // Second column cell – independent content.
            builder.InsertCell();
            // Ensure vertical merge is disabled for this cell.
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row 1, Col 2");

            // End the first row.
            builder.EndRow();

            // ----- Second Row -----
            // First column cell – will be merged with the cell above.
            builder.InsertCell();
            // Mark this cell as a continuation of the previous vertical merge.
            builder.CellFormat.VerticalMerge = CellMerge.Previous;
            builder.Write("This text will be removed after merge");

            // Second column cell – independent content.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row 2, Col 2");

            // End the second row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Retrieve the first column cells from the two rows.
            Cell firstCell = table.Rows[0].Cells[0];
            Cell secondCell = table.Rows[1].Cells[0];

            // Save the document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCells.docx");
            doc.Save(outputPath);

            // Simple verification: the second cell should now be marked as merged (VerticalMerge = Previous).
            bool isMerged = secondCell.CellFormat.VerticalMerge == CellMerge.Previous;
            Console.WriteLine(isMerged
                ? "Cells were merged vertically successfully."
                : "Vertical merge failed.");
        }
    }
}
