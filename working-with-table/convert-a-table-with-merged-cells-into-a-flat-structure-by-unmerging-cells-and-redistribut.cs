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

        // Build a table with a horizontally merged cell (spanning three columns).
        Table table = builder.StartTable();

        // First cell – start of merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged Cell Content");

        // Second cell – part of the merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Third cell – part of the merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Two normal cells after the merged range.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.InsertCell();
        builder.Write("Cell 5");

        builder.EndRow();
        builder.EndTable();

        // Save the original document (optional, just for reference).
        string mergedPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");
        doc.Save(mergedPath);

        // Ensure the table has proper merge flags (convert width‑based merges if any).
        table.ConvertToHorizontallyMergedCells();

        // Unmerge cells and copy the original content into each split cell.
        foreach (Row row in table.Rows)
        {
            for (int i = 0; i < row.Cells.Count; i++)
            {
                Cell cell = row.Cells[i];
                if (cell.CellFormat.HorizontalMerge == CellMerge.First)
                {
                    // Determine how many cells are part of this merged range.
                    int span = 1;
                    int j = i + 1;
                    while (j < row.Cells.Count && row.Cells[j].CellFormat.HorizontalMerge == CellMerge.Previous)
                    {
                        span++;
                        j++;
                    }

                    // Retrieve the text from the original merged cell.
                    string text = cell.GetText().Trim();

                    // Distribute the text to each cell in the span.
                    for (int k = i; k < i + span; k++)
                    {
                        Cell target = row.Cells[k];
                        // Clear existing content.
                        target.RemoveAllChildren();
                        // Add a new paragraph with the text.
                        Paragraph para = new Paragraph(doc);
                        para.AppendChild(new Run(doc, text));
                        target.AppendChild(para);
                        // Reset merge flags.
                        target.CellFormat.HorizontalMerge = CellMerge.None;
                    }

                    // Skip cells that were part of the merged range.
                    i = i + span - 1;
                }
                else if (cell.CellFormat.HorizontalMerge == CellMerge.Previous)
                {
                    // Cells marked as Previous are handled in the First block; just ensure flag is cleared.
                    cell.CellFormat.HorizontalMerge = CellMerge.None;
                }
                else
                {
                    // Non‑merged cells remain unchanged.
                }
            }
        }

        // Save the resulting document with a flat table structure.
        string unmergedPath = Path.Combine(Directory.GetCurrentDirectory(), "UnmergedTable.docx");
        doc.Save(unmergedPath);

        // Simple validation to ensure the output file was created.
        if (!File.Exists(unmergedPath))
            throw new InvalidOperationException("The unmerged table document was not saved correctly.");
    }
}
