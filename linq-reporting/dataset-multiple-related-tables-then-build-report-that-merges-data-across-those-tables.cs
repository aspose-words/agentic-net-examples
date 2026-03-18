using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.MailMerging;

class ReportWithDataSet
{
    static void Main()
    {
        // Create the DataSet with related tables.
        DataSet dataSet = new DataSet();

        // Customers table.
        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("CustomerID", typeof(int));
        customers.Columns.Add("CustomerName", typeof(string));
        customers.Rows.Add(1, "John Doe");
        customers.Rows.Add(2, "Jane Smith");
        dataSet.Tables.Add(customers);

        // Orders table (many‑to‑one relationship with Customers).
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderID", typeof(int));
        orders.Columns.Add("CustomerID", typeof(int));
        orders.Columns.Add("Item", typeof(string));
        orders.Columns.Add("Quantity", typeof(int));
        orders.Rows.Add(101, 1, "Apple", 5);
        orders.Rows.Add(102, 1, "Banana", 3);
        orders.Rows.Add(103, 2, "Orange", 7);
        dataSet.Tables.Add(orders);

        // Define the relation between Customers and Orders.
        dataSet.Relations.Add("CustOrders",
            customers.Columns["CustomerID"],
            orders.Columns["CustomerID"]);

        // Build a template document that contains mail‑merge regions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Outer region for Customers.
        builder.InsertField(" MERGEFIELD TableStart:Customers");
        builder.Writeln("Customer: <<[CustomerName]>>");
        builder.Writeln("Orders:");

        // Inner region for Orders.
        builder.InsertField(" MERGEFIELD TableStart:Orders");
        builder.Writeln("- Item: <<[Item]>>, Qty: <<[Quantity]>>");
        builder.InsertField(" MERGEFIELD TableEnd:Orders");

        // End of the outer region.
        builder.InsertField(" MERGEFIELD TableEnd:Customers");

        // Execute the mail merge using the DataSet with regions.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the resulting report.
        doc.Save("ReportWithDataSet.docx");
    }
}
