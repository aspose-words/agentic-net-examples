using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table.
            Table table = builder.StartTable();

            // ---------- Row 1: horizontally merge three cells ----------
            // First cell – mark as the first cell in a horizontal merge range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Header spanning three columns");

            // Second cell – merge with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Third cell – also merge with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // End the first row.
            builder.EndRow();

            // ---------- Row 2: normal cells ----------
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Row 2, Col 1");

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Row 2, Col 2");

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Row 2, Col 3");

            builder.EndRow();

            // ---------- Row 3: vertically merge first column ----------
            // First cell – start of a vertical merge range.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Vertically merged cell");

            // Second cell – normal.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row 3, Col 2");

            // Third cell – normal.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row 3, Col 3");

            builder.EndRow();

            // ---------- Row 4: continue vertical merge for first column ----------
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous; // Merge with cell above.
            // No text needed for merged cell.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row 4, Col 2");

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row 4, Col 3");

            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document in DOCX format.
            doc.Save("MergedTable.docx");
        }
    }
}
