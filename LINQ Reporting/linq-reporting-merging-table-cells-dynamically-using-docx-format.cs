using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellMergeDemo
{
    // Simple data model for the report.
    public class ReportItem
    {
        public string Category { get; set; }
        public string Product { get; set; }
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Sample data that will be used for the report.
            List<ReportItem> items = new List<ReportItem>
            {
                new ReportItem { Category = "Fruits",   Product = "Apple",  Quantity = 10 },
                new ReportItem { Category = "Fruits",   Product = "Banana", Quantity = 15 },
                new ReportItem { Category = "Fruits",   Product = "Orange", Quantity = 12 },
                new ReportItem { Category = "Vegetables", Product = "Carrot", Quantity = 8 },
                new ReportItem { Category = "Vegetables", Product = "Potato", Quantity = 20 },
                new ReportItem { Category = "Beverages",  Product = "Tea",    Quantity = 5 },
                new ReportItem { Category = "Beverages",  Product = "Coffee", Quantity = 7 }
            };

            // Start building a table with three columns: Category, Product, Quantity.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Category");
            builder.InsertCell();
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Keep a reference to the rows we add so we can later access the cells.
            List<Row> dataRows = new List<Row>();

            // Populate the table with data rows.
            foreach (var item in items)
            {
                builder.InsertCell(); // Category cell (will be merged later)
                builder.Write(item.Category);

                builder.InsertCell(); // Product cell
                builder.Write(item.Product);

                builder.InsertCell(); // Quantity cell
                builder.Write(item.Quantity.ToString());

                builder.EndRow();

                // The builder's current row is the last row added.
                dataRows.Add(table.LastRow);
            }

            // End the table construction.
            builder.EndTable();

            // -----------------------------------------------------------------
            // Dynamically merge Category cells vertically for consecutive rows
            // that share the same Category value.
            // -----------------------------------------------------------------
            // Group rows by Category while preserving the original order.
            var groups = dataRows
                .Select((row, index) => new { Row = row, Index = index })
                .GroupBy(x => items[x.Index].Category)
                .ToList();

            foreach (var group in groups)
            {
                // Get all rows belonging to the current category.
                var rowsInGroup = group.Select(g => g.Row).ToList();

                // Skip groups with a single row – no merge needed.
                if (rowsInGroup.Count < 2) continue;

                // First cell in the group: start of a vertical merge.
                rowsInGroup[0].Cells[0].CellFormat.VerticalMerge = CellMerge.First;

                // Remaining cells: merge with the previous cell.
                for (int i = 1; i < rowsInGroup.Count; i++)
                {
                    rowsInGroup[i].Cells[0].CellFormat.VerticalMerge = CellMerge.Previous;
                }
            }

            // Save the resulting document.
            doc.Save("MergedCellsReport.docx");
        }
    }
}
