using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create a DataSet with related tables.
        DataSet dataSet = new DataSet("SalesData");

        // Customers table
        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("CustomerId", typeof(int));
        customers.Columns.Add("Name", typeof(string));
        customers.Rows.Add(1, "Alice Johnson");
        customers.Rows.Add(2, "Bob Smith");
        dataSet.Tables.Add(customers);

        // Orders table
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderId", typeof(int));
        orders.Columns.Add("CustomerId", typeof(int));
        orders.Columns.Add("OrderDate", typeof(DateTime));
        orders.Rows.Add(1001, 1, new DateTime(2023, 5, 1));
        orders.Rows.Add(1002, 2, new DateTime(2023, 5, 3));
        dataSet.Tables.Add(orders);

        // OrderItems table
        DataTable items = new DataTable("OrderItems");
        items.Columns.Add("ItemId", typeof(int));
        items.Columns.Add("OrderId", typeof(int));
        items.Columns.Add("ProductName", typeof(string));
        items.Columns.Add("Quantity", typeof(int));
        items.Columns.Add("Price", typeof(decimal));
        items.Rows.Add(1, 1001, "Laptop", 1, 1200.00m);
        items.Rows.Add(2, 1001, "Mouse", 2, 25.50m);
        items.Rows.Add(3, 1002, "Desk Chair", 1, 150.00m);
        dataSet.Tables.Add(items);

        // Define relations
        dataSet.Relations.Add("CustomerOrders",
            customers.Columns["CustomerId"],
            orders.Columns["CustomerId"]);

        dataSet.Relations.Add("OrderDetails",
            orders.Columns["OrderId"],
            items.Columns["OrderId"]);

        // 2. Transform DataSet into a strongly‑typed model for reporting.
        ReportModel model = new()
        {
            Orders = new List<OrderReport>()
        };

        foreach (DataRow orderRow in orders.Rows)
        {
            DataRow customerRow = orderRow.GetParentRow("CustomerOrders");
            var orderReport = new OrderReport
            {
                OrderId = (int)orderRow["OrderId"],
                CustomerName = customerRow != null ? (string)customerRow["Name"] : string.Empty,
                OrderDate = ((DateTime)orderRow["OrderDate"]).ToString("yyyy-MM-dd"),
                Items = new List<Item>()
            };

            foreach (DataRow itemRow in orderRow.GetChildRows("OrderDetails"))
            {
                orderReport.Items.Add(new Item
                {
                    ProductName = (string)itemRow["ProductName"],
                    Quantity = (int)itemRow["Quantity"],
                    Price = (decimal)itemRow["Price"]
                });
            }

            model.Orders.Add(orderReport);
        }

        // 3. Create a Word template programmatically with LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "ReportTemplate.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("=== Order Report ===");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Date: <<[order.OrderDate]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.ProductName]>> x <<[item.Quantity]>> @ $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // 4. Load the template and build the report.
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        bool success = engine.BuildReport(reportDoc, model, "model");

        // 5. Save the generated report.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportResult.docx");
        reportDoc.Save(resultPath);

        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to: {resultPath}");
    }
}

// Data model classes
public class ReportModel
{
    public List<OrderReport> Orders { get; set; } = new();
}

public class OrderReport
{
    public int OrderId { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public string OrderDate { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string ProductName { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}
