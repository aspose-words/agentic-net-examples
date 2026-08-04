using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class DetectTableCellFormattingChanges
{
    public static void Main()
    {
        // Prepare a folder for the generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonOutput");
        Directory.CreateDirectory(outputDir);

        // ---------- Create the original document with a simple 2x2 table ----------
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        builder.StartTable();

        // Row 0
        builder.InsertCell();
        builder.Write("A1");
        builder.InsertCell();
        builder.Write("A2");
        builder.EndRow();

        // Row 1
        builder.InsertCell();
        builder.Write("B1");
        builder.InsertCell();
        builder.Write("B2");
        builder.EndRow();

        builder.EndTable();

        string originalPath = Path.Combine(outputDir, "original.docx");
        original.Save(originalPath);

        // ---------- Create the revised document and change the formatting of one cell ----------
        Document revised = (Document)original.Clone(true);

        // Change background color of the cell at row 0, column 1 (second cell of first row)
        Table table = revised.FirstSection.Body.Tables[0];
        Cell targetCell = table.Rows[0].Cells[1];
        targetCell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;

        string revisedPath = Path.Combine(outputDir, "revised.docx");
        revised.Save(revisedPath);

        // ---------- Compare the documents ----------
        // The original document will receive revisions after the comparison.
        original.Compare(revised, "Comparer", DateTime.Now);
        string comparedPath = Path.Combine(outputDir, "compared.docx");
        original.Save(comparedPath);

        // ---------- Inspect revisions and log cell formatting changes ----------
        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange && rev.ParentNode != null)
            {
                // The parent node of a format change on a table cell is a Cell.
                if (rev.ParentNode is Cell changedCell)
                {
                    Row parentRow = changedCell.ParentRow;
                    Table parentTable = parentRow.ParentTable;

                    int rowIndex = parentTable.Rows.IndexOf(parentRow);
                    int columnIndex = parentRow.Cells.IndexOf(changedCell);

                    Console.WriteLine($"Cell format changed at row {rowIndex}, column {columnIndex}");
                }
            }
        }
    }
}
