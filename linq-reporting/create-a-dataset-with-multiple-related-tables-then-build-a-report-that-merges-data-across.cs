using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDatasetExample
{
    public class Program
    {
        public static void Main()
        {
            // 1. Create a DataSet with two related tables: Customers and Orders.
            DataSet dataSet = CreateSampleDataSet();

            // 2. Build a Word template programmatically and insert LINQ Reporting tags.
            const string templatePath = "Template.docx";
            BuildTemplateDocument(templatePath);

            // 3. Load the template and run the LINQ Reporting engine.
            Document report = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;

            // Build the report using the DataSet. An empty string for the data source name
            // allows the template to reference tables directly by their names.
            engine.BuildReport(report, dataSet, "");

            // 4. Save the generated report.
            const string outputPath = "Report.docx";
            report.Save(outputPath);
        }

        // Creates a DataSet containing Customers and Orders tables with a relation.
        private static DataSet CreateSampleDataSet()
        {
            // Customers table.
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("CustomerID", typeof(int));
            customers.Columns.Add("CustomerName", typeof(string));
            customers.Rows.Add(1, "John Doe");
            customers.Rows.Add(2, "Jane Smith");

            // Orders table.
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("CustomerID", typeof(int));
            orders.Columns.Add("OrderName", typeof(string));
            orders.Columns.Add("Quantity", typeof(int));
            orders.Rows.Add(101, 1, "Apples", 5);
            orders.Rows.Add(102, 1, "Bananas", 3);
            orders.Rows.Add(103, 2, "Oranges", 7);

            // Add tables to the DataSet.
            DataSet ds = new DataSet();
            ds.Tables.Add(customers);
            ds.Tables.Add(orders);

            // Create a relation named "CustomerOrders" so that each customer can access its orders.
            DataRelation relation = new DataRelation(
                "CustomerOrders",
                customers.Columns["CustomerID"],
                orders.Columns["CustomerID"]);

            ds.Relations.Add(relation);
            return ds;
        }

        // Generates a simple Word document containing LINQ Reporting tags.
        private static void BuildTemplateDocument(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Begin iterating over customers.
            builder.Writeln("<<foreach [customer in Customers]>>");
            builder.Writeln("Customer: <<[customer.CustomerName]>>");
            builder.Writeln("");

            // Iterate over the related orders for the current customer.
            builder.Writeln("<<foreach [order in customer.CustomerOrders]>>");
            builder.Writeln("  Order: <<[order.OrderName]>> (Qty: <<[order.Quantity]>>)");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            doc.Save(filePath);
        }
    }
}
