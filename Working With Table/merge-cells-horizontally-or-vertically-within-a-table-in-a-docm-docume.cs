using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCM document.
            // Replace "Input.docm" with the path to your source document.
            Document doc = new Document("Input.docm");

            // Ensure the document contains at least one table.
            if (doc.FirstSection.Body.Tables.Count == 0)
                throw new InvalidOperationException("The document does not contain any tables.");

            // Get the first table in the document.
            Table table = doc.FirstSection.Body.Tables[0];

            // -------------------------------------------------
            // Horizontal merge: merge the first two cells of the first row.
            // -------------------------------------------------
            // The leftmost cell becomes the first cell in the merged range.
            table.Rows[0].Cells[0].CellFormat.HorizontalMerge = CellMerge.First;
            // The cell to the right merges with the previous cell.
            table.Rows[0].Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;

            // -------------------------------------------------
            // Vertical merge: merge the first two cells of the first column.
            // -------------------------------------------------
            // The top cell becomes the first cell in the vertically merged range.
            table.Rows[0].Cells[0].CellFormat.VerticalMerge = CellMerge.First;
            // The cell directly below merges with the previous (above) cell.
            // Ensure the table has at least two rows.
            if (table.Rows.Count < 2)
                table.EnsureMinimum(); // Adds a row if missing.
            table.Rows[1].Cells[0].CellFormat.VerticalMerge = CellMerge.Previous;

            // Save the modified document as a DOCM file.
            // Replace "Output.docm" with the desired output path.
            doc.Save("Output.docm");
        }
    }
}
