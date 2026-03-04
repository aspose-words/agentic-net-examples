using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    class PreventCellMerging
    {
        static void Main()
        {
            // Load an existing DOCX document that contains a table.
            Document doc = new Document("InputTable.docx");

            // Iterate through all tables in the document.
            foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
            {
                // Iterate through each row in the current table.
                foreach (Row row in table.Rows)
                {
                    // Iterate through each cell in the current row.
                    foreach (Cell cell in row.Cells)
                    {
                        // Ensure the cell is not part of any horizontal merge range.
                        cell.CellFormat.HorizontalMerge = CellMerge.None;

                        // Ensure the cell is not part of any vertical merge range.
                        cell.CellFormat.VerticalMerge = CellMerge.None;
                    }
                }
            }

            // Save the modified document. The cells will retain their individual formatting
            // and will not be merged even if they contain identical text.
            doc.Save("OutputNoMergedCells.docx");
        }
    }
}
