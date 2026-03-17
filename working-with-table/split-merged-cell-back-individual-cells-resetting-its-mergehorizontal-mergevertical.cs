using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitMergedCells
{
    static void Main()
    {
        // Load the document that contains merged cells.
        Document doc = new Document("MergedCells.docx");

        // Iterate through every table in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Iterate through each row and each cell.
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    // Reset both horizontal and vertical merge flags to None.
                    cell.CellFormat.HorizontalMerge = CellMerge.None;
                    cell.CellFormat.VerticalMerge = CellMerge.None;
                }
            }
        }

        // Save the document with the cells split back into individual cells.
        doc.Save("UnmergedCells.docx");
    }
}
