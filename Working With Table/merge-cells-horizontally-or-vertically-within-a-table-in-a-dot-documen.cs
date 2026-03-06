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

            // Use DocumentBuilder to construct the table.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -------------------------------------------------
            // Build a table with 2 rows and 3 columns.
            // -------------------------------------------------
            builder.StartTable();

            // ---------- Row 1 ----------
            // Cell 1 – start of a horizontally merged range (spans two columns).
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First; // first cell in the merge group
            builder.Write("Horizontally merged cells");

            // Cell 2 – merges with the previous cell horizontally.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous; // continues the merge

            // Cell 3 – a normal, unmerged cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None; // reset merge flag for subsequent cells
            builder.Write("Normal cell");

            // End the first row.
            builder.EndRow();

            // ---------- Row 2 ----------
            // Cell 1 – normal cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None; // ensure no horizontal merge
            builder.Write("Row2, Col1");

            // Cell 2 – start of a vertically merged range (spans two rows).
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First; // first cell in the vertical merge group
            builder.Write("Vertically merged cells");

            // Cell 3 – normal cell.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None; // reset vertical merge flag
            builder.Write("Row2, Col3");

            // End the second row.
            builder.EndRow();

            // ---------- Row 3 ----------
            // Cell 1 – normal cell.
            builder.InsertCell();
            builder.Write("Row3, Col1");

            // Cell 2 – merges with the cell above vertically.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous; // continues vertical merge
            // No text needed for merged cells; they must be empty.
            
            // Cell 3 – normal cell.
            builder.InsertCell();
            builder.Write("Row3, Col3");

            // End the third row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to disk.
            string outputPath = @"C:\Temp\MergedTable.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
