using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Sample data that will be used for the report.
        var data = new[]
        {
            new { Category = "Fruits",      Name = "Apple"  },
            new { Category = "Fruits",      Name = "Banana" },
            new { Category = "Vegetables", Name = "Carrot" },
            new { Category = "Vegetables", Name = "Lettuce"},
            new { Category = "Vegetables", Name = "Tomato" }
        };

        // Group the data by Category – each group will occupy a set of rows
        // where the Category cell is merged horizontally across those rows.
        var groups = data.GroupBy(d => d.Category);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table that will hold the report.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Category");
        builder.InsertCell();
        builder.Write("Item");
        builder.EndRow();

        // Iterate over each group and write its rows.
        foreach (var grp in groups)
        {
            bool firstRowInGroup = true;

            foreach (var item in grp)
            {
                // First column – will be merged horizontally for the whole group.
                builder.InsertCell();

                if (firstRowInGroup)
                {
                    // Mark this cell as the first cell in a merged range.
                    builder.CellFormat.HorizontalMerge = CellMerge.First;
                    builder.Write(grp.Key); // Write the category name only once.
                    firstRowInGroup = false;
                }
                else
                {
                    // Subsequent cells in the same group merge with the previous cell.
                    builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                }

                // Second column – normal (no merge) cell with the item name.
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.None;
                builder.Write(item.Name);

                // End the current row.
                builder.EndRow();
            }
        }

        // Finish the table.
        builder.EndTable();

        // Save the document in DOCX format.
        doc.Save("MergedCellsReport.docx");
    }
}
