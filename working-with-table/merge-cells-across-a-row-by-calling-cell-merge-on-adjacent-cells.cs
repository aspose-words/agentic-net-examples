using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMergeDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add three cells in the first row.
            builder.StartTable();

            // First cell – will be the first cell in a merged range.
            Cell firstCell = builder.InsertCell();
            builder.Write("Cell 1");
            firstCell.CellFormat.HorizontalMerge = CellMerge.First;

            // Second cell – will be merged with the first cell.
            Cell secondCell = builder.InsertCell();
            builder.Write("Cell 2 (will be merged)");
            secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Third cell – remains independent.
            Cell thirdCell = builder.InsertCell();
            builder.Write("Cell 3");
            thirdCell.CellFormat.HorizontalMerge = CellMerge.None;

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Optional validation: the second cell should now have HorizontalMerge = CellMerge.Previous.
            if (secondCell.CellFormat.HorizontalMerge != CellMerge.Previous)
                throw new InvalidOperationException("Cell merging failed.");

            // Save the document to the local file system.
            string outputPath = "MergedCells.docx";
            doc.Save(outputPath);

            // Inform the user that the operation completed.
            Console.WriteLine($"Document saved to '{outputPath}'.");
        }
    }
}
