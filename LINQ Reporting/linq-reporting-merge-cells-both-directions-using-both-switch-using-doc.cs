using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellMergeExample
{
    // Enum to describe the type of merge for a cell.
    enum MergeDirection
    {
        None,
        HorizontalFirst,
        HorizontalPrevious,
        VerticalFirst,
        VerticalPrevious
    }

    // Simple data model representing a cell and its merge direction.
    class CellInfo
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public MergeDirection Direction { get; set; }
        public string Text { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define a collection of cells with merge instructions.
            // This collection will be processed with LINQ to build the table.
            List<CellInfo> cells = new List<CellInfo>
            {
                // First row – merge first two cells horizontally.
                new CellInfo { Row = 0, Column = 0, Direction = MergeDirection.HorizontalFirst, Text = "Header 1-2" },
                new CellInfo { Row = 0, Column = 1, Direction = MergeDirection.HorizontalPrevious, Text = "" },
                new CellInfo { Row = 0, Column = 2, Direction = MergeDirection.None, Text = "Header 3" },

                // Second row – merge first column vertically with the cell above.
                new CellInfo { Row = 1, Column = 0, Direction = MergeDirection.VerticalPrevious, Text = "" },
                new CellInfo { Row = 1, Column = 1, Direction = MergeDirection.None, Text = "R2C2" },
                new CellInfo { Row = 1, Column = 2, Direction = MergeDirection.None, Text = "R2C3" },

                // Third row – normal cells.
                new CellInfo { Row = 2, Column = 0, Direction = MergeDirection.None, Text = "R3C1" },
                new CellInfo { Row = 2, Column = 1, Direction = MergeDirection.None, Text = "R3C2" },
                new CellInfo { Row = 2, Column = 2, Direction = MergeDirection.None, Text = "R3C3" }
            };

            // Start building the table.
            builder.StartTable();

            // Group cells by row to handle row creation.
            var rows = cells.GroupBy(c => c.Row).OrderBy(g => g.Key);
            foreach (var rowGroup in rows)
            {
                // Process each cell in the current row ordered by column index.
                foreach (var cellInfo in rowGroup.OrderBy(c => c.Column))
                {
                    // Insert a new cell.
                    builder.InsertCell();

                    // Apply merge settings using a switch statement.
                    switch (cellInfo.Direction)
                    {
                        case MergeDirection.HorizontalFirst:
                            builder.CellFormat.HorizontalMerge = CellMerge.First;
                            break;
                        case MergeDirection.HorizontalPrevious:
                            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                            break;
                        case MergeDirection.VerticalFirst:
                            builder.CellFormat.VerticalMerge = CellMerge.First;
                            break;
                        case MergeDirection.VerticalPrevious:
                            builder.CellFormat.VerticalMerge = CellMerge.Previous;
                            break;
                        case MergeDirection.None:
                        default:
                            // Ensure no merge flags are set for normal cells.
                            builder.CellFormat.HorizontalMerge = CellMerge.None;
                            builder.CellFormat.VerticalMerge = CellMerge.None;
                            break;
                    }

                    // Write the cell text if any.
                    if (!string.IsNullOrEmpty(cellInfo.Text))
                        builder.Write(cellInfo.Text);
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document in DOC format.
            doc.Save("MergedCellsReport.doc", SaveFormat.Doc);
        }
    }
}
