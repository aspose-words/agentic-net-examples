using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace MergeHeaderCellsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the source document that contains tables with a header row.
            // The header row is expected to contain the placeholder text "<<mergeCells>>".
            Document doc = new Document("Input.docx");

            // Iterate through all tables in the document.
            foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
            {
                // Ensure the table has at least one row.
                if (table.Rows.Count == 0)
                    continue;

                // Work with the first row (header row).
                Row headerRow = table.FirstRow;

                // Find the first cell that contains the merge placeholder.
                int placeholderIndex = -1;
                for (int i = 0; i < headerRow.Cells.Count; i++)
                {
                    Cell cell = headerRow.Cells[i];
                    string cellText = cell.ToString(SaveFormat.Text).Trim();
                    if (cellText.Contains("<<mergeCells>>"))
                    {
                        placeholderIndex = i;
                        break;
                    }
                }

                // If no placeholder was found, skip this table.
                if (placeholderIndex == -1)
                    continue;

                // Determine how many cells should be merged.
                // For this example we merge all cells from the placeholder cell to the end of the row.
                int cellsToMerge = headerRow.Cells.Count - placeholderIndex;

                // Apply horizontal merge flags.
                for (int i = 0; i < cellsToMerge; i++)
                {
                    Cell cell = headerRow.Cells[placeholderIndex + i];
                    if (i == 0)
                        cell.CellFormat.HorizontalMerge = CellMerge.First;      // First cell in the merged range.
                    else
                        cell.CellFormat.HorizontalMerge = CellMerge.Previous; // Merge with the previous cell.
                }

                // Optionally, clear the placeholder text from the first merged cell.
                Cell firstMergedCell = headerRow.Cells[placeholderIndex];
                firstMergedCell.RemoveAllChildren(); // Remove existing paragraphs.
                firstMergedCell.AppendChild(new Paragraph(doc));
                firstMergedCell.FirstParagraph.AppendChild(new Run(doc, "Merged Header"));
            }

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
