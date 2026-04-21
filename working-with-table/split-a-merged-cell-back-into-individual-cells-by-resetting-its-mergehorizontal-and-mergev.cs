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

            // Second cell – part of the merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text needed for merged cells.
            builder.EndRow();

            // Add a normal second row.
            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndTable();

            // Split the merged cell back into individual cells by resetting merge flags.
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.HorizontalMerge = CellMerge.None;
                    cell.CellFormat.VerticalMerge = CellMerge.None;
                }
            }

            // Save the resulting document.
            string outputPath = "SplitMergedCell.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to {outputPath}");
        }
    }
}
