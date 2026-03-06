using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellMergePrevent
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCX document (replace with your actual path)
            string inputPath = @"Input.docx";
            Document doc = new Document(inputPath);

            // Iterate through all tables in the document
            foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
            {
                // Process each row
                for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
                {
                    Row row = table.Rows[rowIndex];

                    // Process each cell in the current row
                    for (int colIndex = 0; colIndex < row.Cells.Count; colIndex++)
                    {
                        Cell cell = row.Cells[colIndex];
                        string cellText = cell.GetText().Trim();

                        // Check left neighbour for identical text (horizontal)
                        if (colIndex > 0)
                        {
                            Cell leftCell = row.Cells[colIndex - 1];
                            string leftText = leftCell.GetText().Trim();

                            if (string.Equals(cellText, leftText, StringComparison.Ordinal))
                            {
                                // Prevent horizontal merging by setting both cells to None
                                cell.CellFormat.HorizontalMerge = CellMerge.None;
                                leftCell.CellFormat.HorizontalMerge = CellMerge.None;
                            }
                        }

                        // Check upper neighbour for identical text (vertical)
                        if (rowIndex > 0)
                        {
                            Row upperRow = table.Rows[rowIndex - 1];
                            // Ensure the upper row has enough cells (tables can be irregular)
                            if (colIndex < upperRow.Cells.Count)
                            {
                                Cell upperCell = upperRow.Cells[colIndex];
                                string upperText = upperCell.GetText().Trim();

                                if (string.Equals(cellText, upperText, StringComparison.Ordinal))
                                {
                                    // Prevent vertical merging by setting both cells to None
                                    cell.CellFormat.VerticalMerge = CellMerge.None;
                                    upperCell.CellFormat.VerticalMerge = CellMerge.None;
                                }
                            }
                        }

                        // Explicitly set merge type to None for safety
                        cell.CellFormat.HorizontalMerge = CellMerge.None;
                        cell.CellFormat.VerticalMerge = CellMerge.None;
                    }
                }
            }

            // Save the modified document (replace with your desired output path)
            string outputPath = @"Output.docx";
            doc.Save(outputPath);
        }
    }
}
