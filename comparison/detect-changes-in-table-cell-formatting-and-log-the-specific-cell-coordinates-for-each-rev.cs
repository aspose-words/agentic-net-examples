using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableCellFormattingComparison
{
    public static void Main()
    {
        // Create the original document with a simple 2x2 table.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.StartTable();
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row}C{col}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Create the revised document with the same content.
        Document revisedDoc = new Document();
        DocumentBuilder revBuilder = new DocumentBuilder(revisedDoc);
        revBuilder.StartTable();
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                revBuilder.InsertCell();
                revBuilder.Writeln($"R{row}C{col}");
            }
            revBuilder.EndRow();
        }
        revBuilder.EndTable();

        // Apply formatting changes to specific cells in the revised document.
        Table revisedTable = revisedDoc.FirstSection.Body.Tables[0];

        // Change background color of cell at (0,1).
        Cell cell01 = revisedTable.Rows[0].Cells[1];
        cell01.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;

        // Change text alignment of cell at (1,0) by aligning its paragraph.
        Cell cell10 = revisedTable.Rows[1].Cells[0];
        if (cell10.Paragraphs.Count > 0)
        {
            cell10.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
        }

        // Compare the documents.
        originalDoc.Compare(revisedDoc, "Comparer", DateTime.Now);

        // Log cell formatting revisions with their coordinates.
        foreach (Revision revision in originalDoc.Revisions)
        {
            if (revision.RevisionType == RevisionType.FormatChange &&
                revision.ParentNode?.NodeType == NodeType.Cell)
            {
                Cell changedCell = (Cell)revision.ParentNode!;
                Row parentRow = changedCell.ParentRow;
                Table parentTable = parentRow.ParentTable;

                int rowIndex = parentTable.Rows.IndexOf(parentRow);
                int columnIndex = parentRow.Cells.IndexOf(changedCell);

                Console.WriteLine($"Cell formatting revision at Row {rowIndex}, Column {columnIndex}");
            }
        }

        // Save the compared document.
        originalDoc.Save("comparison_result.docx");
    }
}
