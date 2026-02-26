using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsMhtmlMerge
{
    class Program
    {
        static void Main()
        {
            // Load an existing MHTML document.
            // This follows the provided "load" rule: new Document(string path).
            string inputPath = "input.mht";
            Document doc = new Document(inputPath);

            // Ensure the document contains at least one table.
            if (doc.FirstSection?.Body?.Tables?.Count > 0)
            {
                // Get the first table in the document.
                Table table = doc.FirstSection.Body.Tables[0];

                // Verify that the table has at least two rows and two columns
                // to demonstrate both horizontal and vertical merging.
                if (table.Rows.Count >= 2 && table.Rows[0].Cells.Count >= 2)
                {
                    // ---------- Horizontal merge (first row) ----------
                    // The first cell becomes the start of a horizontally merged range.
                    table.Rows[0].Cells[0].CellFormat.HorizontalMerge = CellMerge.First;
                    // The second cell merges with the cell to its left.
                    table.Rows[0].Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;

                    // ---------- Vertical merge (first column) ----------
                    // The first cell becomes the start of a vertically merged range.
                    table.Rows[0].Cells[0].CellFormat.VerticalMerge = CellMerge.First;
                    // The cell directly below merges with the cell above it.
                    table.Rows[1].Cells[0].CellFormat.VerticalMerge = CellMerge.Previous;
                }
            }

            // Save the modified document back to MHTML format.
            // This follows the provided "save" rule: doc.Save(string path, SaveFormat format).
            string outputPath = "output.mht";
            doc.Save(outputPath, SaveFormat.Mhtml);
        }
    }
}
