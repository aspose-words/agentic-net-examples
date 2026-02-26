using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    class PreventIdenticalCellMerging
    {
        static void Main()
        {
            // Load an existing DOC document that contains a table.
            // The document may have cells merged automatically because they contain identical text.
            Document doc = new Document(@"Input.doc");

            // Iterate through all tables in the document.
            foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
            {
                // Iterate through each row of the current table.
                foreach (Row row in table.Rows)
                {
                    // Iterate through each cell of the current row.
                    foreach (Cell cell in row.Cells)
                    {
                        // Explicitly disable any horizontal or vertical merging.
                        // This prevents Aspose.Words (and Word) from treating the cell as part of a merged range.
                        cell.CellFormat.HorizontalMerge = CellMerge.None;
                        cell.CellFormat.VerticalMerge   = CellMerge.None;
                    }
                }
            }

            // Save the modified document in DOC format.
            doc.Save(@"Output.doc", SaveFormat.Doc);
        }
    }
}
