using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDataSetExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Customer
    {
        public int CustomerId { get; set; }
        public string Name { get; set; } = "";
        public List<Order> Orders { get; set; } = new();
    }

    public class Order
    {
        public int OrderId { get; set; }
        public DateTime OrderDate { get; set; }
        public List<OrderDetail> Details { get; set; } = new();
    }

    public class OrderDetail
    {
        public int OrderDetailId { get; set; }
        public string Product { get; set; } = "";
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output directory
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a DataSet with related tables
            DataSet dataSet = CreateSampleDataSet();

            // 2. Transform DataSet into hierarchical model for reporting
            ReportModel model = BuildReportModel(dataSet);

            // 3. Create the template document programmatically
            string templatePath = Path.Combine(outputDir, "Template.docx");
            CreateTemplateDocument(templatePath);

            // 4. Load the template and build the report
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // 5. Save the generated report
            string reportPath = Path.Combine(outputDir, "Report.docx");
            reportDoc.Save(reportPath);
        }

        private static DataSet CreateSampleDataSet()
        {
            DataSet ds = new DataSet("ShopData");

            // Customers table
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("CustomerId", typeof(int));
            customers.Columns.Add("Name", typeof(string));
            customers.Rows.Add(1, "Alice");
            customers.Rows.Add(2, "Bob");
            ds.Tables.Add(customers);

            // Orders table
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderId", typeof(int));
            orders.Columns.Add("CustomerId", typeof(int));
            orders.Columns.Add("OrderDate", typeof(DateTime));
            orders.Rows.Add(100, 1, new DateTime(2023, 1, 15));
            orders.Rows.Add(101, 1, new DateTime(2023, 2, 5));
            orders.Rows.Add(102, 2, new DateTime(2023, 3, 12));
            ds.Tables.Add(orders);

            // OrderDetails table
            DataTable details = new DataTable("OrderDetails");
            details.Columns.Add("OrderDetailId", typeof(int));
            details.Columns.Add("OrderId", typeof(int));
            details.Columns.Add("Product", typeof(string));
            details.Columns.Add("Quantity", typeof(int));
            details.Rows.Add(1000, 100, "Laptop", 1);
            details.Rows.Add(1001, 100, "Mouse", 2);
            details.Rows.Add(1002, 101, "Keyboard", 1);
            details.Rows.Add(1003, 102, "Monitor", 2);
            ds.Tables.Add(details);

            // Define relationships (optional, not required for LINQ)
            ds.Relations.Add("Customer_Orders",
                ds.Tables["Customers"].Columns["CustomerId"],
                ds.Tables["Orders"].Columns["CustomerId"]);
            ds.Relations.Add("Order_Details",
                ds.Tables["Orders"].Columns["OrderId"],
                ds.Tables["OrderDetails"].Columns["OrderId"]);

            return ds;
        }

        private static ReportModel BuildReportModel(DataSet ds)
        {
            var customers = ds.Tables["Customers"]
                .AsEnumerable()
                .Select(c => new Customer
                {
                    CustomerId = c.Field<int>("CustomerId"),
                    Name = c.Field<string>("Name") ?? "",
                    Orders = ds.Tables["Orders"]
                        .AsEnumerable()
                        .Where(o => o.Field<int>("CustomerId") == c.Field<int>("CustomerId"))
                        .Select(o => new Order
                        {
                            OrderId = o.Field<int>("OrderId"),
                            OrderDate = o.Field<DateTime>("OrderDate"),
                            Details = ds.Tables["OrderDetails"]
                                .AsEnumerable()
                                .Where(d => d.Field<int>("OrderId") == o.Field<int>("OrderId"))
                                .Select(d => new OrderDetail
                                {
                                    OrderDetailId = d.Field<int>("OrderDetailId"),
                                    Product = d.Field<string>("Product") ?? "",
                                    Quantity = d.Field<int>("Quantity")
                                })
                                .ToList()
                        })
                        .ToList()
                })
                .ToList();

            return new ReportModel { Customers = customers };
        }

        private static void CreateTemplateDocument(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Customer Orders Report");
            builder.Writeln();

            // Begin foreach over customers
            builder.Writeln("<<foreach [cust in Customers]>>");
            builder.Writeln("Customer: <<[cust.Name]>>");
            builder.Writeln();

            // Begin foreach over orders of the current customer
            builder.Writeln("<<foreach [ord in cust.Orders]>>");
            builder.Writeln("  Order ID: <<[ord.OrderId]>>   Date: <<[ord.OrderDate]>>");
            builder.Writeln();

            // Begin foreach over details of the current order
            builder.Writeln("  <<foreach [det in ord.Details]>>");
            builder.Writeln("    Product: <<[det.Product]>>   Qty: <<[det.Quantity]>>");
            builder.Writeln("  <</foreach>>");
            builder.Writeln();

            // End orders foreach
            builder.Writeln("<</foreach>>");
            builder.Writeln();

            // End customers foreach
            builder.Writeln("<</foreach>>");

            doc.Save(filePath);
        }
    }
}
