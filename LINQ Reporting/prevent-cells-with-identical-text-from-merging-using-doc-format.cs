using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class PreventCellMerging
{
    static void Main()
    {
        // Load an existing DOCX (or any supported) document that contains tables.
        Document doc = new Document("Input.docx");

        // Iterate through all tables in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Iterate through each cell in the current table.
            foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
            {
                // Ensure the cell is not part of any horizontal or vertical merge.
                // This prevents Word from automatically merging cells that have identical text
                // when the document is saved in the legacy DOC format.
                cell.CellFormat.HorizontalMerge = CellMerge.None;
                cell.CellFormat.VerticalMerge   = CellMerge.None;
            }
        }

        // Save the modified document in the legacy DOC format.
        // The SaveOptions object can be used to fine‑tune the output if needed.
        doc.Save("Output.doc", SaveFormat.Doc);
    }
}
