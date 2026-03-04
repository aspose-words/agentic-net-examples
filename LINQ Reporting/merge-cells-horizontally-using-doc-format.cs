using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellMergeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder which will be used to construct the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            builder.StartTable();

            // Insert the first cell of the first row.
            // This cell will be the first cell in a range of horizontally merged cells.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged across two cells");

            // Insert the second cell of the first row.
            // This cell will be merged with the previous cell horizontally.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // End the first row.
            builder.EndRow();

            // Insert two more cells in the second row that are not merged.
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.Write("Cell 1, Row 2");

            builder.InsertCell();
            builder.Write("Cell 2, Row 2");

            // End the second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to disk.
            doc.Save("CellFormat.HorizontalMerge.docx");
        }
    }
}
