using System;
using Aspose.Words;
using Aspose.Words.Tables;

class DetectCellFormatRevisions
{
    static void Main()
    {
        // Load the document (lifecycle rule: load)
        Document doc = new Document("InputWithRevisions.docx");

        // Iterate through all revisions in the document
        foreach (Revision rev in doc.Revisions)
        {
            // Look for format revisions whose parent node is a table cell
            if (rev.RevisionType == RevisionType.FormatChange && rev.ParentNode?.NodeType == NodeType.Cell)
            {
                Cell cell = (Cell)rev.ParentNode;
                Row row = cell.ParentRow;
                Table table = row.ParentTable;

                // Calculate 1‑based row and column indices
                int rowIndex = table.IndexOf(row) + 1;
                int columnIndex = row.IndexOf(cell) + 1;

                // Log the cell coordinates that contain a formatting revision
                Console.WriteLine($"Format revision detected in cell: Row {rowIndex}, Column {columnIndex}");
            }
        }

        // Save the document (lifecycle rule: save)
        doc.Save("OutputWithRevisionsLogged.docx");
    }
}
