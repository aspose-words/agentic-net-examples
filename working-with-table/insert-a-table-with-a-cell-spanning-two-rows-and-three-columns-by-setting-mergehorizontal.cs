using System;
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

            // Start a new table.
            Table table = builder.StartTable();

            // ---------- First row ----------
            // Cell (0,0) – top‑left cell that will span 2 rows and 3 columns.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Spanning 2 rows x 3 columns");

            // Cell (0,1) – part of the horizontal merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge = CellMerge.None;

            // Cell (0,2) – part of the horizontal merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge = CellMerge.None;

            // End the first row.
            builder.EndRow();

            // ---------- Second row ----------
            // Cell (1,0) – part of the vertical merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.VerticalMerge = CellMerge.Previous;

            // Cell (1,1) – part of the horizontal merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge = CellMerge.None;

            // Cell (1,2) – part of the horizontal merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge = CellMerge.None;

            // End the second row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document.
            const string outputPath = "MergedTable.docx";
            doc.Save(outputPath);
        }
    }
}
