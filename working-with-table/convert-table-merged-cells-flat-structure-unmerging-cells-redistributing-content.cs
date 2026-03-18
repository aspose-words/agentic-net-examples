using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source document that contains tables with merged cells.
        Document doc = new Document("Input.docx");

        // Iterate through all tables in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Convert cells that were merged by width into cells merged by flags.
            // This makes the merge information explicit and easier to manipulate.
            table.ConvertToHorizontallyMergedCells();

            // Unmerge all cells by clearing both horizontal and vertical merge flags.
            // After this operation each cell becomes independent.
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    // Reset horizontal merge flag.
                    cell.CellFormat.HorizontalMerge = CellMerge.None;

                    // Reset vertical merge flag (if any vertical merges exist).
                    cell.CellFormat.VerticalMerge = CellMerge.None;
                }
            }
        }

        // Save the modified document with a flat table structure.
        doc.Save("Output.docx");
    }
}
