using System;
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

            // First cell – start of a horizontally merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged cell content");

            // Second cell – merge with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text is required for the merged cell.

            // End the first row.
            builder.EndRow();

            // Add a second row with regular (unmerged) cells for demonstration.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Row 2, Cell 1");

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Row 2, Cell 2");

            // End the second row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document.
            const string outputPath = "MergedCells.docx";
            doc.Save(outputPath);
        }
    }
}
