using System;
using System.IO;
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

            // Ensure the output directory exists.
            string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Start a 3x3 table.
            builder.StartTable();

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    // Insert a new cell.
                    builder.InsertCell();

                    // Determine merge settings based on cell position.
                    if (row == 0 && col == 0)
                    {
                        // This cell is the first cell of a horizontal and vertical merge range.
                        builder.CellFormat.HorizontalMerge = CellMerge.First;
                        builder.CellFormat.VerticalMerge   = CellMerge.First;
                        builder.Write("Merged Cell");
                    }
                    else if (row == 0 && col == 1)
                    {
                        // Merge this cell horizontally with the cell to its left.
                        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                        builder.CellFormat.VerticalMerge   = CellMerge.None;
                    }
                    else if (row == 1 && col == 0)
                    {
                        // Merge this cell vertically with the cell above.
                        builder.CellFormat.HorizontalMerge = CellMerge.None;
                        builder.CellFormat.VerticalMerge   = CellMerge.Previous;
                    }
                    else
                    {
                        // Regular, unmerged cell.
                        builder.CellFormat.HorizontalMerge = CellMerge.None;
                        builder.CellFormat.VerticalMerge   = CellMerge.None;
                        builder.Write($"R{row + 1}C{col + 1}");
                    }
                }

                // End the current row.
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Save the document in DOCX format.
            string outputPath = Path.Combine(artifactsDir, "MergedTable.docx");
            doc.Save(outputPath);
        }
    }
}
