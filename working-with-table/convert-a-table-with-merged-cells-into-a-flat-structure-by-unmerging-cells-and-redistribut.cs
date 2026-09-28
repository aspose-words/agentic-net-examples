using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with merged cells.
        // Row 0: [Merged horizontally (col 0-1)] | [Cell 2 (merged vertically)]
        // Row 1: [Cell 0] | [Cell 1] | [Cell 2 (merged vertically)]
        builder.StartTable();

        // First row, first cell (will be merged horizontally with next cell)
        builder.InsertCell();
        builder.Writeln("Merged H");

        // First row, second cell (merged horizontally)
        builder.InsertCell();
        // Set horizontal merge flag to indicate this cell continues the merge.
        ((Cell)builder.CurrentParagraph.ParentNode).CellFormat.HorizontalMerge = CellMerge.Previous;

        // First row, third cell (will be merged vertically)
        builder.InsertCell();
        builder.Writeln("Merged V");
        // Set vertical merge flag for the first cell in the vertical merge.
        ((Cell)builder.CurrentParagraph.ParentNode).CellFormat.VerticalMerge = CellMerge.First;

        builder.EndRow();

        // Second row
        // Cell 0
        builder.InsertCell();
        builder.Writeln("Cell 0");

        // Cell 1
        builder.InsertCell();
        builder.Writeln("Cell 1");

        // Cell 2 (continuation of vertical merge)
        builder.InsertCell();
        // Set vertical merge flag to indicate continuation.
        ((Cell)builder.CurrentParagraph.ParentNode).CellFormat.VerticalMerge = CellMerge.Previous;

        builder.EndRow();

        builder.EndTable();

        // Save the document with merged cells.
        string mergedPath = "MergedTable.docx";
        doc.Save(mergedPath);

        // Load the document for processing.
        Document processedDoc = new Document(mergedPath);
        Table table = (Table)processedDoc.GetChild(NodeType.Table, 0, true);

        // ----- Unmerge horizontally -----
        for (int rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
        {
            Row row = table.Rows[rowIdx];
            for (int cellIdx = 0; cellIdx < row.Cells.Count; cellIdx++)
            {
                Cell cell = row.Cells[cellIdx];
                if (cell.CellFormat.HorizontalMerge == CellMerge.First)
                {
                    // Capture the original text.
                    string originalText = cell.GetText().Trim();

                    // Unmerge the first cell.
                    cell.CellFormat.HorizontalMerge = CellMerge.None;

                    // Propagate text to all cells that were merged horizontally.
                    int nextIdx = cellIdx + 1;
                    while (nextIdx < row.Cells.Count &&
                           row.Cells[nextIdx].CellFormat.HorizontalMerge == CellMerge.Previous)
                    {
                        Cell mergedCell = row.Cells[nextIdx];
                        mergedCell.CellFormat.HorizontalMerge = CellMerge.None;

                        // Clear existing content and add the original text.
                        mergedCell.RemoveAllChildren();
                        Paragraph para = new Paragraph(processedDoc);
                        mergedCell.AppendChild(para);
                        para.AppendChild(new Run(processedDoc, originalText));

                        nextIdx++;
                    }
                }
            }
        }

        // ----- Unmerge vertically -----
        // Determine the maximum number of columns in the table.
        int maxColumns = 0;
        foreach (Row r in table.Rows)
            if (r.Cells.Count > maxColumns) maxColumns = r.Cells.Count;

        for (int colIdx = 0; colIdx < maxColumns; colIdx++)
        {
            for (int rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
            {
                // Guard against rows that may have fewer cells than the current column index.
                if (colIdx >= table.Rows[rowIdx].Cells.Count) continue;

                Cell cell = table.Rows[rowIdx].Cells[colIdx];
                if (cell.CellFormat.VerticalMerge == CellMerge.First)
                {
                    string originalText = cell.GetText().Trim();
                    cell.CellFormat.VerticalMerge = CellMerge.None;

                    int nextRow = rowIdx + 1;
                    while (nextRow < table.Rows.Count)
                    {
                        if (colIdx >= table.Rows[nextRow].Cells.Count) break;

                        Cell mergedCell = table.Rows[nextRow].Cells[colIdx];
                        if (mergedCell.CellFormat.VerticalMerge == CellMerge.Previous)
                        {
                            mergedCell.CellFormat.VerticalMerge = CellMerge.None;
                            mergedCell.RemoveAllChildren();
                            Paragraph para = new Paragraph(processedDoc);
                            mergedCell.AppendChild(para);
                            para.AppendChild(new Run(processedDoc, originalText));
                            nextRow++;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
        }

        // Save the flattened table.
        string flatPath = "FlatTable.docx";
        processedDoc.Save(flatPath);

        // Simple validation that the output file exists.
        if (!File.Exists(flatPath))
            throw new Exception("The flattened table document was not created.");
    }
}
