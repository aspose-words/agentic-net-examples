using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a merged table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // First row, first cell (will be merged horizontally and vertically).
        builder.InsertCell();
        builder.Write("Merged Cell");

        // First row, second cell (merged horizontally with the first cell).
        builder.InsertCell();
        // Set horizontal merge to "Previous" on this cell.
        Cell secondCellRow1 = (Cell)builder.CurrentParagraph.ParentNode;
        secondCellRow1.CellFormat.HorizontalMerge = CellMerge.Previous;

        // First row, third cell (stand‑alone).
        builder.InsertCell();
        builder.Write("A1");

        // End first row.
        builder.EndRow();

        // Second row, first cell (merged vertically with the first cell of the first row).
        builder.InsertCell();
        // Set vertical merge to "Previous" on this cell.
        Cell firstCellRow2 = (Cell)builder.CurrentParagraph.ParentNode;
        firstCellRow2.CellFormat.VerticalMerge = CellMerge.Previous;

        // Second row, second cell (part of the same horizontal merge region).
        builder.InsertCell();
        // Set horizontal merge to "Previous" on this cell.
        Cell secondCellRow2 = (Cell)builder.CurrentParagraph.ParentNode;
        secondCellRow2.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Second row, third cell (stand‑alone).
        builder.InsertCell();
        builder.Write("B2");

        // End second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document containing the merged table.
        const string mergedPath = "MergedTable.docx";
        doc.Save(mergedPath);

        // Load the document again to perform unmerge processing.
        Document loadedDoc = new Document(mergedPath);

        // Process each table in the document.
        foreach (Node node in loadedDoc.GetChildNodes(NodeType.Table, true))
        {
            Table table = (Table)node;

            // Iterate through rows and cells.
            for (int rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
            {
                Row row = table.Rows[rowIdx];
                for (int cellIdx = 0; cellIdx < row.Cells.Count; cellIdx++)
                {
                    Cell cell = row.Cells[cellIdx];

                    bool isHorizontalContinuation = cell.CellFormat.HorizontalMerge == CellMerge.Previous;
                    bool isVerticalContinuation = cell.CellFormat.VerticalMerge == CellMerge.Previous;

                    if (isHorizontalContinuation || isVerticalContinuation)
                    {
                        // Find the origin (top‑left) cell of the merged region.
                        int originRowIdx = rowIdx;
                        int originCellIdx = cellIdx;

                        // Move left while previous cells are marked as horizontal continuation.
                        while (originCellIdx > 0 &&
                               table.Rows[originRowIdx].Cells[originCellIdx].CellFormat.HorizontalMerge == CellMerge.Previous)
                        {
                            originCellIdx--;
                        }

                        // Move up while previous rows are marked as vertical continuation.
                        while (originRowIdx > 0 &&
                               table.Rows[originRowIdx].Cells[originCellIdx].CellFormat.VerticalMerge == CellMerge.Previous)
                        {
                            originRowIdx--;
                        }

                        Cell originCell = table.Rows[originRowIdx].Cells[originCellIdx];

                        // Copy all child nodes from the origin cell to the current cell.
                        foreach (Node child in originCell.GetChildNodes(NodeType.Any, true))
                        {
                            cell.AppendChild(child.Clone(true));
                        }

                        // Reset merge flags for the current cell.
                        cell.CellFormat.HorizontalMerge = CellMerge.None;
                        cell.CellFormat.VerticalMerge = CellMerge.None;
                    }
                }
            }

            // After processing all cells, reset merge flags on all cells (including original origins).
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.HorizontalMerge = CellMerge.None;
                    cell.CellFormat.VerticalMerge = CellMerge.None;
                }
            }
        }

        // Save the flattened table document.
        const string flatPath = "FlatTable.docx";
        loadedDoc.Save(flatPath);

        // Simple validation that the output file exists.
        if (!File.Exists(flatPath))
            throw new InvalidOperationException($"Failed to create the output file '{flatPath}'.");
    }
}
