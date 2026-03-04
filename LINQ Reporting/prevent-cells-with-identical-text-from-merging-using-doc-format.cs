using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace PreventCellMerging
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the source DOC document.
            Document doc = new Document("Input.doc");

            // Iterate through all tables in the document.
            foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
            {
                // Iterate through each row of the current table.
                foreach (Row row in table.Rows)
                {
                    // Iterate through each cell of the current row.
                    foreach (Cell cell in row.Cells)
                    {
                        // Explicitly set both horizontal and vertical merge types to None.
                        // This ensures that cells with identical text are not automatically merged
                        // when the document is saved in DOC format.
                        cell.CellFormat.HorizontalMerge = CellMerge.None;
                        cell.CellFormat.VerticalMerge = CellMerge.None;
                    }
                }
            }

            // Save the modified document. The cells will retain their individual boundaries.
            doc.Save("Output.doc");
        }
    }
}
