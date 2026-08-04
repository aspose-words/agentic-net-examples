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
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a table with a merged cell (both horizontally and vertically).
            // The table will have 2 rows and 2 columns.
            Table table = builder.StartTable();

            // First row, first cell – start of a merged region.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;   // Horizontal merge start.
            builder.CellFormat.VerticalMerge = CellMerge.First;     // Vertical merge start.
            builder.Write("Merged Cell");

            // First row, second cell – merge with the cell to the left.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Horizontal merge continuation.
            builder.CellFormat.VerticalMerge = CellMerge.None;       // No vertical merge.
            builder.Write(string.Empty); // Empty content for merged cell.

            builder.EndRow();

            // Second row, first cell – merge vertically with the cell above.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;     // No horizontal merge.
            builder.CellFormat.VerticalMerge = CellMerge.Previous;   // Vertical merge continuation.
            builder.Write(string.Empty); // Empty content for merged cell.

            // Second row, second cell – normal unmerged cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Normal Cell");

            builder.EndRow();
            builder.EndTable();

            // At this point the table contains a merged cell.
            // Now split the merged cell back into individual cells by resetting merge properties.
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.HorizontalMerge = CellMerge.None;
                    cell.CellFormat.VerticalMerge = CellMerge.None;
                }
            }

            // Save the document to the local file system.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "SplitMergedCell.docx");
            doc.Save(outputPath);
        }
    }
}
