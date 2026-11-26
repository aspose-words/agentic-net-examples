using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a merged-cell table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // First row: merge first two cells horizontally.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged Header");

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged part.

        builder.EndRow();

        // Second row: normal cells.
        builder.InsertCell();
        builder.Write("Row 1, Col 1");

        builder.InsertCell();
        builder.Write("Row 1, Col 2");

        builder.EndRow();

        // Third row: merge first cell vertically with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically Merged Cell");

        builder.InsertCell();
        builder.Write("Row 2, Col 2");

        builder.EndRow();

        // Fourth row: continuation of vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed.

        builder.InsertCell();
        builder.Write("Row 3, Col 2");

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the original document (optional, for reference).
        string originalPath = "OriginalTable.docx";
        doc.Save(originalPath);

        // Process the document to unmerge cells and redistribute content.
        UnmergeTableCells(doc);

        // Save the processed document.
        string outputPath = "UnmergedTable.docx";
        doc.Save(outputPath);

        // Validate that the output file exists.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate success (no interactive output required).
        Console.WriteLine("Table unmerged and saved to " + outputPath);
    }

    private static void UnmergeTableCells(Document doc)
    {
        // Iterate through all tables in the document.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        foreach (Table table in tables)
        {
            // Build a 2D list of cells for easy navigation.
            List<List<Cell>> cellMatrix = new List<List<Cell>>();
            foreach (Row row in table.Rows)
            {
                List<Cell> cellRow = new List<Cell>();
                foreach (Cell cell in row.Cells)
                {
                    cellRow.Add(cell);
                }
                cellMatrix.Add(cellRow);
            }

            // First pass: capture master cell text for each merged region.
            // Dictionary maps a cell (by its reference) to the master text.
            Dictionary<Cell, string> masterTextMap = new Dictionary<Cell, string>();

            for (int rowIdx = 0; rowIdx < cellMatrix.Count; rowIdx++)
            {
                List<Cell> row = cellMatrix[rowIdx];
                for (int colIdx = 0; colIdx < row.Count; colIdx++)
                {
                    Cell cell = row[colIdx];
                    // Determine if this cell is the start (master) of a merged region.
                    bool isHorizontalFirst = cell.CellFormat.HorizontalMerge == CellMerge.First;
                    bool isVerticalFirst = cell.CellFormat.VerticalMerge == CellMerge.First;

                    if (isHorizontalFirst || isVerticalFirst || (cell.CellFormat.HorizontalMerge == CellMerge.None && cell.CellFormat.VerticalMerge == CellMerge.None))
                    {
                        // This cell is a master for its region (or a normal cell).
                        string text = GetCellText(cell);
                        masterTextMap[cell] = text;
                    }
                }
            }

            // Second pass: assign each cell the text of its master cell.
            for (int rowIdx = 0; rowIdx < cellMatrix.Count; rowIdx++)
            {
                List<Cell> row = cellMatrix[rowIdx];
                for (int colIdx = 0; colIdx < row.Count; colIdx++)
                {
                    Cell cell = row[colIdx];

                    // Find master cell coordinates.
                    int masterRow = rowIdx;
                    int masterCol = colIdx;

                    // Move left while this cell is a horizontal continuation.
                    while (cell.CellFormat.HorizontalMerge == CellMerge.Previous && masterCol > 0)
                    {
                        masterCol--;
                        cell = row[masterCol];
                    }

                    // Reset cell reference to original for vertical search.
                    cell = row[colIdx];

                    // Move up while this cell is a vertical continuation.
                    while (cell.CellFormat.VerticalMerge == CellMerge.Previous && masterRow > 0)
                    {
                        masterRow--;
                        cell = cellMatrix[masterRow][colIdx];
                    }

                    Cell masterCell = cellMatrix[masterRow][masterCol];
                    string masterText = masterTextMap.ContainsKey(masterCell) ? masterTextMap[masterCell] : "";

                    // Reset merges for the current cell.
                    Cell currentCell = cellMatrix[rowIdx][colIdx];
                    currentCell.CellFormat.HorizontalMerge = CellMerge.None;
                    currentCell.CellFormat.VerticalMerge = CellMerge.None;

                    // Set the cell's text to the master text.
                    SetCellText(currentCell, masterText);
                }
            }
        }
    }

    private static string GetCellText(Cell cell)
    {
        // Retrieve the concatenated text of all paragraphs in the cell.
        return cell.GetText().TrimEnd('\a'); // Remove cell end marker.
    }

    private static void SetCellText(Cell cell, string text)
    {
        // Clear existing content.
        cell.RemoveAllChildren();

        // Add a new paragraph with the specified text.
        Paragraph para = new Paragraph(cell.Document);
        Run run = new Run(cell.Document, text);
        para.AppendChild(run);
        cell.AppendChild(para);
    }
}
