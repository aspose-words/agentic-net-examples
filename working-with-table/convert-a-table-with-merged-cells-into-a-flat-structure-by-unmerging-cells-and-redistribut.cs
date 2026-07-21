using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableUnmergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with a table that contains horizontally merged cells.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start the table.
            Table table = builder.StartTable();

            // First row – merge first three cells horizontally.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged Cell");

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Add a normal cell.
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.Write("Normal Cell");

            builder.EndRow();

            // Second row – normal cells only.
            builder.InsertCell();
            builder.Write("Row2 Cell1");
            builder.InsertCell();
            builder.Write("Row2 Cell2");
            builder.EndRow();

            builder.EndTable();

            // Save the original document (optional, for reference).
            doc.Save("OriginalTable.docx");

            // -----------------------------------------------------------------
            // Flatten the table: unmerge cells and duplicate the original content
            // into each cell that was part of a merged range.
            // -----------------------------------------------------------------
            // Get the first table in the document.
            Table originalTable = doc.FirstSection.Body.Tables[0];

            // Create a new table that will hold the flattened structure.
            Table flatTable = new Table(doc);
            doc.FirstSection.Body.InsertAfter(flatTable, originalTable);

            // Process each row of the original table.
            foreach (Row originalRow in originalTable.Rows)
            {
                Row newRow = new Row(doc);
                flatTable.AppendChild(newRow);

                int cellIndex = 0;
                while (cellIndex < originalRow.Cells.Count)
                {
                    Cell currentCell = originalRow.Cells[cellIndex];
                    CellMerge hMerge = currentCell.CellFormat.HorizontalMerge;

                    if (hMerge == CellMerge.First)
                    {
                        // Determine how many cells are merged together.
                        int mergedCount = 1;
                        int nextIndex = cellIndex + 1;
                        while (nextIndex < originalRow.Cells.Count &&
                               originalRow.Cells[nextIndex].CellFormat.HorizontalMerge == CellMerge.Previous)
                        {
                            mergedCount++;
                            nextIndex++;
                        }

                        // Get the text from the first cell of the merged range.
                        string cellText = currentCell.GetText().Trim();

                        // Create separate cells for each part of the merged range,
                        // duplicating the original text.
                        for (int i = 0; i < mergedCount; i++)
                        {
                            Cell newCell = new Cell(doc);
                            newCell.AppendChild(new Paragraph(doc));
                            newCell.FirstParagraph.AppendChild(new Run(doc, cellText));
                            newRow.AppendChild(newCell);
                        }

                        // Move the index past the merged range.
                        cellIndex = nextIndex;
                    }
                    else if (hMerge == CellMerge.Previous)
                    {
                        // This cell is part of a merged range that has already been handled.
                        cellIndex++;
                    }
                    else // CellMerge.None – normal cell.
                    {
                        // Clone the cell (including its content) and add it to the new row.
                        Cell clonedCell = (Cell)currentCell.Clone(true);
                        newRow.AppendChild(clonedCell);
                        cellIndex++;
                    }
                }
            }

            // Remove the original table from the document.
            originalTable.Remove();

            // Save the resulting document with the flattened table.
            doc.Save("FlattenedTable.docx");
        }
    }
}
