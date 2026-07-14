using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

namespace TableUnmergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with a table that contains merged cells.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start the table.
            Table table = builder.StartTable();

            // First row – horizontally merged cells (cells 1‑2) and a normal cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged H1");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous; // part of the merge
            builder.Write(""); // content will be redistributed later
            builder.InsertCell(); // third cell, not merged
            builder.Write("Normal");
            builder.EndRow();

            // Second row – vertically merged cells (cells 1‑2) and a normal cell.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Merged V1");
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous; // part of the vertical merge
            builder.Write(""); // content will be redistributed later
            builder.InsertCell();
            builder.Write("Normal2");
            builder.EndRow();

            // Third row – normal cells.
            builder.InsertCell();
            builder.Write("Cell3-1");
            builder.InsertCell();
            builder.Write("Cell3-2");
            builder.InsertCell();
            builder.Write("Cell3-3");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the original document (optional, for reference).
            doc.Save("OriginalTable.docx");

            // -----------------------------------------------------------------
            // Unmerge the table and redistribute the merged content.
            // -----------------------------------------------------------------
            // Get the first table in the document.
            Table targetTable = doc.FirstSection.Body.Tables[0];

            // First, handle horizontal merges.
            foreach (Row row in targetTable.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    Cell cell = row.Cells[i];
                    if (cell.CellFormat.HorizontalMerge == CellMerge.First)
                    {
                        // Capture the text from the first cell of the merge.
                        string mergedText = cell.GetText().Trim();

                        // Propagate the text to all subsequent cells that are marked as Previous.
                        int j = i + 1;
                        while (j < row.Cells.Count && row.Cells[j].CellFormat.HorizontalMerge == CellMerge.Previous)
                        {
                            Cell prevCell = row.Cells[j];
                            // Clear existing content.
                            prevCell.RemoveAllChildren();
                            // Add a new paragraph with the same text.
                            Paragraph para = new Paragraph(doc);
                            para.AppendChild(new Run(doc, mergedText));
                            prevCell.AppendChild(para);
                            j++;
                        }
                    }
                }
            }

            // Next, handle vertical merges.
            // Determine the maximum number of columns in the table.
            int maxCols = 0;
            foreach (Row row in targetTable.Rows)
                if (row.Cells.Count > maxCols) maxCols = row.Cells.Count;

            for (int col = 0; col < maxCols; col++)
            {
                for (int rowIdx = 0; rowIdx < targetTable.Rows.Count; rowIdx++)
                {
                    Row row = targetTable.Rows[rowIdx];
                    if (col >= row.Cells.Count) continue; // safety check

                    Cell cell = row.Cells[col];
                    if (cell.CellFormat.VerticalMerge == CellMerge.First)
                    {
                        // Capture the text from the first cell of the vertical merge.
                        string mergedText = cell.GetText().Trim();

                        // Propagate the text to all subsequent cells in the same column marked as Previous.
                        int nextRow = rowIdx + 1;
                        while (nextRow < targetTable.Rows.Count)
                        {
                            Row nextRowObj = targetTable.Rows[nextRow];
                            if (col >= nextRowObj.Cells.Count) break;
                            Cell nextCell = nextRowObj.Cells[col];
                            if (nextCell.CellFormat.VerticalMerge != CellMerge.Previous) break;

                            // Clear existing content.
                            nextCell.RemoveAllChildren();
                            // Add a new paragraph with the same text.
                            Paragraph para = new Paragraph(doc);
                            para.AppendChild(new Run(doc, mergedText));
                            nextCell.AppendChild(para);

                            nextRow++;
                        }
                    }
                }
            }

            // Finally, remove all merge flags so the table becomes a flat structure.
            foreach (Row row in targetTable.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.HorizontalMerge = CellMerge.None;
                    cell.CellFormat.VerticalMerge = CellMerge.None;
                }
            }

            // Save the resulting document.
            doc.Save("UnmergedTable.docx");
        }
    }
}
