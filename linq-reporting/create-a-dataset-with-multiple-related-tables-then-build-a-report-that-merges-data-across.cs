using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a DataSet with two related tables: Customers and Orders.
        DataSet dataSet = new DataSet();

        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("CustomerID", typeof(int));
        customers.Columns.Add("CustomerName", typeof(string));
        customers.Rows.Add(1, "John Doe");
        customers.Rows.Add(2, "Jane Smith");
        dataSet.Tables.Add(customers);

        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderID", typeof(int));
        orders.Columns.Add("CustomerID", typeof(int));
        orders.Columns.Add("ProductName", typeof(string));
        orders.Columns.Add("Quantity", typeof(int));
        orders.Rows.Add(1001, 1, "Laptop", 1);
        orders.Rows.Add(1002, 1, "Mouse", 2);
        orders.Rows.Add(1003, 2, "Keyboard", 1);
        dataSet.Tables.Add(orders);

        // Define a relation between Customers and Orders on CustomerID.
        dataSet.Relations.Add("CustomerOrders",
            customers.Columns["CustomerID"],
            orders.Columns["CustomerID"]);

        // Build a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Customers and Orders Report");
        builder.Writeln();

        // Outer loop over customers.
        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln("Customer: <<[c.CustomerName]>> (ID: <<[c.CustomerID]>>)");
        builder.Writeln();

        // Inner loop over related orders using the defined relation.
        builder.Writeln("Orders:");
        builder.Writeln("<<foreach [o in c.CustomerOrders]>>");
        builder.Writeln("- Order <<[o.OrderID]>>: <<[o.ProductName]>> (Qty: <<[o.Quantity]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options required.
        bool success = engine.BuildReport(doc, dataSet, "");

        // Save the generated report.
        doc.Save("Report.docx");

        // Indicate success (optional).
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}
