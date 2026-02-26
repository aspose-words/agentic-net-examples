using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample data source.
        var items = new List<Item>
        {
            new Item { Category = "Fruits",      Name = "Apple"  },
            new Item { Category = "Fruits",      Name = "Banana" },
            new Item { Category = "Vegetables",  Name = "Carrot" },
            new Item { Category = "Vegetables",  Name = "Lettuce"},
            new Item { Category = "Vegetables",  Name = "Pepper" }
        };

        // Start a table with two columns: Category and Name.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Category");
        builder.InsertCell();
        builder.Write("Name");
        builder.EndRow();

        // Group items by Category and merge the Category cells vertically.
        var groups = items.GroupBy(i => i.Category);
        foreach (var group in groups)
        {
            bool firstRowInGroup = true;
            foreach (var item in group)
            {
                // Category cell.
                builder.InsertCell();
                if (firstRowInGroup)
                {
                    // First cell in the vertical merge range.
                    builder.Write(group.Key);
                    builder.CellFormat.VerticalMerge = CellMerge.First;
                }
                else
                {
                    // Merge with the cell above.
                    builder.Write(string.Empty);
                    builder.CellFormat.VerticalMerge = CellMerge.Previous;
                }

                // Name cell.
                builder.InsertCell();
                builder.Write(item.Name);
                // Reset vertical merge for the Name column (no merging needed).
                builder.CellFormat.VerticalMerge = CellMerge.None;

                builder.EndRow();

                // Prepare for the next iteration.
                builder.CellFormat.HorizontalMerge = CellMerge.None;
                firstRowInGroup = false;
            }
        }

        builder.EndTable();

        // Save the resulting document.
        doc.Save("MergedTable.docx");
    }

    // Simple data model used for the LINQ grouping.
    class Item
    {
        public string Category { get; set; }
        public string Name { get; set; }
    }
}
