using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a DataSet with related tables.
        DataSet dataSet = new DataSet();

        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("CustomerId", typeof(int));
        customers.Columns.Add("CustomerName", typeof(string));
        customers.Rows.Add(1, "Alice");
        customers.Rows.Add(2, "Bob");
        dataSet.Tables.Add(customers);

        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderId", typeof(int));
        orders.Columns.Add("CustomerId", typeof(int));
        orders.Columns.Add("OrderDate", typeof(DateTime));
        orders.Rows.Add(100, 1, new DateTime(2023, 1, 10));
        orders.Rows.Add(101, 1, new DateTime(2023, 2, 5));
        orders.Rows.Add(102, 2, new DateTime(2023, 3, 12));
        dataSet.Tables.Add(orders);

        DataTable orderDetails = new DataTable("OrderDetails");
        orderDetails.Columns.Add("OrderDetailId", typeof(int));
        orderDetails.Columns.Add("OrderId", typeof(int));
        orderDetails.Columns.Add("Product", typeof(string));
        orderDetails.Columns.Add("Quantity", typeof(int));
        orderDetails.Rows.Add(1000, 100, "Laptop", 1);
        orderDetails.Rows.Add(1001, 100, "Mouse", 2);
        orderDetails.Rows.Add(1002, 101, "Keyboard", 1);
        orderDetails.Rows.Add(1003, 102, "Monitor", 2);
        dataSet.Tables.Add(orderDetails);

        // Build a hierarchical model from the DataSet.
        List<Customer> customerList = (from c in customers.AsEnumerable()
            select new Customer
            {
                Name = c.Field<string>("CustomerName"),
                Orders = (from o in orders.AsEnumerable()
                    where o.Field<int>("CustomerId") == c.Field<int>("CustomerId")
                    select new Order
                    {
                        OrderId = o.Field<int>("OrderId"),
                        Details = (from d in orderDetails.AsEnumerable()
                            where d.Field<int>("OrderId") == o.Field<int>("OrderId")
                            select new OrderDetail
                            {
                                Product = d.Field<string>("Product"),
                                Quantity = d.Field<int>("Quantity")
                            }).ToList()
                    }).ToList()
            }).ToList();

        ReportModel model = new ReportModel { Customers = customerList };

        // Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Customer Orders Report");
        builder.Writeln("<<foreach [c in model.Customers]>>");
        builder.Writeln("Customer: <<[c.Name]>>");
        builder.Writeln("Orders:");
        builder.Writeln("<<foreach [o in c.Orders]>>");
        builder.Writeln("- Order ID: <<[o.OrderId]>>");
        builder.Writeln("  Details:");
        builder.Writeln("  <<foreach [d in o.Details]>>");
        builder.Writeln("    * <<[d.Product]>> (Qty: <<[d.Quantity]>>)");
        builder.Writeln("  <</foreach>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template (demonstrating load before building report).
        Document doc = new Document(templatePath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Root model class.
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

// Customer class.
public class Customer
{
    public string Name { get; set; } = "";
    public List<Order> Orders { get; set; } = new();
}

// Order class.
public class Order
{
    public int OrderId { get; set; }
    public List<OrderDetail> Details { get; set; } = new();
}

// Order detail class.
public class OrderDetail
{
    public string Product { get; set; } = "";
    public int Quantity { get; set; }
}
