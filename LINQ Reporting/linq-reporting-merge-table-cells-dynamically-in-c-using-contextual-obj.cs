using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Reporting;

namespace AsposeWordsDynamicCellMerge
{
    // Simple data model.
    public class Customer
    {
        public string Name { get; set; }
    }

    public class Order
    {
        public Customer Customer { get; set; }
        public string Product { get; set; }
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data.
            var orders = new List<Order>
            {
                new Order { Customer = new Customer { Name = "Alice" }, Product = "Apple",  Quantity = 5 },
                new Order { Customer = new Customer { Name = "Alice" }, Product = "Banana", Quantity = 3 },
                new Order { Customer = new Customer { Name = "Bob" },   Product = "Carrot", Quantity = 7 },
                new Order { Customer = new Customer { Name = "Bob" },   Product = "Date",   Quantity = 2 },
                new Order { Customer = new Customer { Name = "Bob" },   Product = "Eggplant", Quantity = 4 },
                new Order { Customer = new Customer { Name = "Carol" }, Product = "Fig",    Quantity = 6 }
            };

            // Group orders by customer name – this will drive the vertical merge.
            var groups = orders
                .GroupBy(o => o.Customer.Name)
                .Select(g => new
                {
                    CustomerName = g.Key,
                    Orders = g.ToList()
                })
                .ToList();

            // Create a blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a table with three columns: Customer, Product, Quantity.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Customer");
            builder.InsertCell();
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Iterate over each group and create rows.
            foreach (var group in groups)
            {
                // Number of rows needed for this group.
                int rowCount = group.Orders.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    // Insert cells for the current row.
                    // Customer column – merge vertically across the group.
                    builder.InsertCell();
                    if (i == 0)
                    {
                        // First cell in the merged range.
                        builder.CellFormat.VerticalMerge = CellMerge.First;
                        builder.Write(group.CustomerName);
                    }
                    else
                    {
                        // Subsequent cells merge with the previous one.
                        builder.CellFormat.VerticalMerge = CellMerge.Previous;
                        // No text needed for merged cells.
                    }

                    // Reset vertical merge for the next columns to avoid unintended merging.
                    builder.CellFormat.VerticalMerge = CellMerge.None;

                    // Product column.
                    builder.InsertCell();
                    builder.Write(group.Orders[i].Product);

                    // Quantity column.
                    builder.InsertCell();
                    builder.Write(group.Orders[i].Quantity.ToString());

                    builder.EndRow();
                }
            }

            // End the table.
            builder.EndTable();

            // OPTIONAL: Demonstrate ReportingEngine with DOT notation.
            // Create a simple template that uses contextual object member access.
            Document template = new Document();
            DocumentBuilder tmplBuilder = new DocumentBuilder(template);
            tmplBuilder.Writeln("<<foreach [order]>>");
            tmplBuilder.Writeln("Customer: <<[order.Customer.Name]>>");
            tmplBuilder.Writeln("Product : <<[order.Product]>>");
            tmplBuilder.Writeln("Qty     : <<[order.Quantity]>>");
            tmplBuilder.Writeln("<<endfor>>");

            // Build the report using the same orders list as the data source.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "order" matches the placeholder used in the template.
            engine.BuildReport(template, orders, "order");

            // Save both documents.
            doc.Save("DynamicCellMerge.docx");
            template.Save("ReportingEngineTemplate.docx");
        }
    }
}
