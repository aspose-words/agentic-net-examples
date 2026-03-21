using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class DetectCellFormatRevisions
{
    static void Main()
    {
        // Create a new document and add a simple table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Start tracking revisions.
        doc.StartTrackRevisions("User");

        // Change the formatting of a cell to generate a format revision.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell targetCell = table.Rows[0].Cells[1];
        targetCell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Iterate through all revisions in the document.
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange && rev.ParentNode?.NodeType == NodeType.Cell)
            {
                Cell cell = (Cell)rev.ParentNode;
                Row row = cell.ParentRow;
                Table tbl = row.ParentTable;

                int rowIndex = tbl.IndexOf(row) + 1;
                int columnIndex = row.IndexOf(cell) + 1;

                Console.WriteLine($"Format revision detected in cell: Row {rowIndex}, Column {columnIndex}");
            }
        }

        // Save the document.
        doc.Save("OutputWithRevisionsLogged.docx");
    }
}
