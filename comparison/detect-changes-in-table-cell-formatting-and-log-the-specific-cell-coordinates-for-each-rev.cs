using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a deterministic output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // ---------- Create the original document ----------
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        // Build a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        string originalPath = Path.Combine(outputDir, "Original.docx");
        original.Save(originalPath);

        // ---------- Clone and modify formatting ----------
        Document edited = (Document)original.Clone(true);
        Table? table = edited.GetChild(NodeType.Table, 0, true) as Table;
        if (table != null)
        {
            // Change background of first cell.
            Cell cell11 = table.Rows[0].Cells[0];
            cell11.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;

            // Change background of second row, second column.
            Cell cell22 = table.Rows[1].Cells[1];
            cell22.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        }

        string editedPath = Path.Combine(outputDir, "Edited.docx");
        edited.Save(editedPath);

        // ---------- Compare documents ----------
        // The comparison adds format revisions to the original document.
        original.Compare(edited, "Comparer", DateTime.Now);
        string comparedPath = Path.Combine(outputDir, "Compared.docx");
        original.Save(comparedPath);

        // ---------- Inspect revisions for cell format changes ----------
        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange && rev.ParentNode.NodeType == NodeType.Cell)
            {
                Cell changedCell = (Cell)rev.ParentNode;
                Row? parentRow = changedCell.ParentRow;
                Table? parentTable = parentRow?.ParentTable;

                int rowIndex = parentTable?.Rows.IndexOf(parentRow) ?? -1;
                int columnIndex = parentRow?.Cells.IndexOf(changedCell) ?? -1;

                // Output coordinates using 1‑based indexing.
                Console.WriteLine($"Format change detected at Row {rowIndex + 1}, Column {columnIndex + 1}");
            }
        }
    }
}
