using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // DocumentBuilder simplifies building the document structure.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Define the size of the table.
        int rows = 3;
        int cols = 4;

        // Build the table row by row.
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                // Insert a new cell at the current cursor position.
                builder.InsertCell();

                // ----- Horizontal merge example -----
                // Merge the first two cells of the first row.
                if (r == 0 && c == 0)
                {
                    // Mark this cell as the first cell in a horizontally merged range.
                    builder.CellFormat.HorizontalMerge = CellMerge.First;
                    builder.Write("Header");
                }
                else if (r == 0 && c == 1)
                {
                    // Mark this cell as merged to the previous cell.
                    builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                }
                // ----- Vertical merge example -----
                // Merge the third column (index 2) vertically across all rows.
                else if (c == 2)
                {
                    if (r == 0)
                    {
                        // First cell in a vertically merged range.
                        builder.CellFormat.VerticalMerge = CellMerge.First;
                        builder.Write("Vertical");
                    }
                    else
                    {
                        // Subsequent cells merged to the cell above.
                        builder.CellFormat.VerticalMerge = CellMerge.Previous;
                    }
                }
                else
                {
                    // For all other cells, ensure no merging flags are set.
                    builder.CellFormat.HorizontalMerge = CellMerge.None;
                    builder.CellFormat.VerticalMerge = CellMerge.None;
                    builder.Write($"R{r + 1}C{c + 1}");
                }
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("MergedCells.docx");
    }
}
