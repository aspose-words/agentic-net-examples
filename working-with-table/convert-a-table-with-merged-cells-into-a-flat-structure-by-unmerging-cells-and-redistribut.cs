using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with merged cells.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Start a table.
        builder.StartTable();

        // First row – horizontally merged cells (2 cells merged).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Horizontally merged text");
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // The merged cell does not need its own text.
        builder.EndRow();

        // Second row – vertically merged cells (first column merged across two rows).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged text");
        builder.InsertCell();
        builder.Write("Normal cell 2, row 2");
        builder.EndRow();

        // Third row – continuation of the vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed; will be filled after unmerging.
        builder.InsertCell();
        builder.Write("Normal cell 2, row 3");
        builder.EndRow();

        builder.EndTable();

        // Save the source document.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");
        sourceDoc.Save(sourcePath);

        // Load the document for processing.
        Document doc = new Document(sourcePath);
        Table table = doc.FirstSection.Body.Tables[0];

        // Ensure merged cells are represented by merge flags.
        table.ConvertToHorizontallyMergedCells();

        // ---------- Unmerge horizontally ----------
        foreach (Row row in table.Rows)
        {
            for (int col = 0; col < row.Cells.Count; col++)
            {
                Cell cell = row.Cells[col];
                if (cell.CellFormat.HorizontalMerge == CellMerge.First)
                {
                    // Determine how many cells are part of this horizontal merge.
                    int mergeCount = 1;
                    int nextCol = col + 1;
                    while (nextCol < row.Cells.Count &&
                           row.Cells[nextCol].CellFormat.HorizontalMerge == CellMerge.Previous)
                    {
                        mergeCount++;
                        nextCol++;
                    }

                    // Capture the original text.
                    string text = cell.GetText().Trim();

                    // Distribute the text to each cell in the merged range and clear merge flags.
                    for (int i = 0; i < mergeCount; i++)
                    {
                        Cell target = row.Cells[col + i];
                        target.RemoveAllChildren();
                        Paragraph para = new Paragraph(doc);
                        para.AppendChild(new Run(doc, text));
                        target.AppendChild(para);
                        target.CellFormat.HorizontalMerge = CellMerge.None;
                    }

                    // Skip the cells we have already processed.
                    col += mergeCount - 1;
                }
            }
        }

        // ---------- Unmerge vertically ----------
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            Row row = table.Rows[rowIndex];
            for (int col = 0; col < row.Cells.Count; col++)
            {
                Cell cell = row.Cells[col];
                if (cell.CellFormat.VerticalMerge == CellMerge.First)
                {
                    // Determine how many rows are part of this vertical merge.
                    int mergeRows = 1;
                    int nextRow = rowIndex + 1;
                    while (nextRow < table.Rows.Count &&
                           table.Rows[nextRow].Cells.Count > col &&
                           table.Rows[nextRow].Cells[col].CellFormat.VerticalMerge == CellMerge.Previous)
                    {
                        mergeRows++;
                        nextRow++;
                    }

                    // Capture the original text.
                    string text = cell.GetText().Trim();

                    // Distribute the text to each cell in the merged column range and clear merge flags.
                    for (int i = 0; i < mergeRows; i++)
                    {
                        Cell target = table.Rows[rowIndex + i].Cells[col];
                        target.RemoveAllChildren();
                        Paragraph para = new Paragraph(doc);
                        para.AppendChild(new Run(doc, text));
                        target.AppendChild(para);
                        target.CellFormat.VerticalMerge = CellMerge.None;
                    }
                }
            }
        }

        // Validate that no cells remain merged.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                if (cell.CellFormat.HorizontalMerge != CellMerge.None ||
                    cell.CellFormat.VerticalMerge != CellMerge.None)
                {
                    throw new Exception("Unmerge operation failed: some cells are still merged.");
                }
            }
        }

        // Save the flattened table.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FlatTable.docx");
        doc.Save(outputPath);
    }
}
