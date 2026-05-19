using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableSplit
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a table with a horizontally merged cell in the first row.
            Table table = builder.StartTable();

            // First cell – start of the merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged cell");

            // Second cell – merged to the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // End the first row.
            builder.EndRow();

            // Second row – normal unmerged cells.
            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // -------- Split the merged cell --------
            // Reset merge flags for every cell in the table.
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.HorizontalMerge = CellMerge.None;
                    cell.CellFormat.VerticalMerge = CellMerge.None;
                }
            }

            // Save the document to verify the result.
            string outputPath = "SplitMergedCell.docx";
            doc.Save(outputPath);
        }
    }
}
