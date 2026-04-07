using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDataSetExample
{
    public class Program
    {
        public static void Main()
        {
            // 1. Create a DataSet with related tables.
            DataSet dataSet = CreateDataSet();

            // 2. Build a Word template programmatically.
            const string templatePath = "Template.docx";
            BuildTemplate(templatePath);

            // 3. Load the template (required before building the report).
            Document template = new Document(templatePath);

            // 4. Build the report using the DataSet.
            ReportingEngine engine = new ReportingEngine();
            // Allow missing members so that any typo does not abort the build.
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            engine.BuildReport(template, dataSet, "ds");

            // 5. Save the generated report.
            const string reportPath = "Report.docx";
            template.Save(reportPath);

            Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
        }

        // Creates a DataSet containing Customers and Orders tables with a relation.
        private static DataSet CreateDataSet()
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
            orders.Columns.Add("ProductName", typeof(string));
            orders.Columns.Add("Quantity", typeof(int));
            orders.Rows.Add(101, 1, "Apple", 5);
            orders.Rows.Add(102, 1, "Banana", 3);
            orders.Rows.Add(103, 2, "Orange", 7);

            // Create the DataSet and add tables.
            DataSet ds = new DataSet();
            ds.Tables.Add(customers);
            ds.Tables.Add(orders);

            // Define a relation so that each Customer row can access its Orders.
            DataRelation relation = new DataRelation(
                "Orders",                     // Relation name (used as property name in LINQ Reporting)
                customers.Columns["CustomerID"],
                orders.Columns["CustomerID"]);
            ds.Relations.Add(relation);

            return ds;
        }

        // Generates a Word document that contains LINQ Reporting tags.
        private static void BuildTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Title.
            builder.Writeln("Customer Orders Report");
            builder.Writeln();

            // Outer loop: iterate over customers.
            builder.Writeln("<<foreach [cust in ds.Customers]>>");
            builder.Writeln("Customer: <<[cust.CustomerName]>>");
            builder.Writeln();

            // Inner loop: iterate over orders belonging to the current customer.
            builder.Writeln("Orders:");
            builder.Writeln("<<foreach [ord in ds.Orders]>>");
            builder.Writeln("<<if [ord.CustomerID == cust.CustomerID]>>- <<[ord.ProductName]>> (Qty: <<[ord.Quantity]>>)<</if>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            doc.Save(filePath);
        }
    }
}
