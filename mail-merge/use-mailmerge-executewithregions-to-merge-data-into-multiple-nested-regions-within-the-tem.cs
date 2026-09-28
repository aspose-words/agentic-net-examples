using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeNestedRegionsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a DataSet with two related tables: Customers and Orders.
            DataSet dataSet = new DataSet();

            // Customers table.
            DataTable customersTable = new DataTable("Customers");
            customersTable.Columns.Add("Name", typeof(string));
            customersTable.Columns.Add("Country", typeof(string));

            customersTable.Rows.Add("Alice", "USA");
            customersTable.Rows.Add("Bob", "Canada");
            dataSet.Tables.Add(customersTable);

            // Orders table.
            DataTable ordersTable = new DataTable("Orders");
            ordersTable.Columns.Add("CustomerName", typeof(string)); // Foreign key to Customers.Name
            ordersTable.Columns.Add("OrderID", typeof(int));
            ordersTable.Columns.Add("Product", typeof(string));
            ordersTable.Columns.Add("Quantity", typeof(int));

            ordersTable.Rows.Add("Alice", 1001, "Laptop", 1);
            ordersTable.Rows.Add("Alice", 1002, "Mouse", 2);
            ordersTable.Rows.Add("Bob", 2001, "Keyboard", 1);
            dataSet.Tables.Add(ordersTable);

            // Define relation between Customers and Orders.
            DataRelation relation = new DataRelation(
                "CustomerOrders",
                customersTable.Columns["Name"],
                ordersTable.Columns["CustomerName"]);
            dataSet.Relations.Add(relation);

            // Build a template document with nested mail‑merge regions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Customer Report");
            builder.Writeln("<<TableStart:Customers>>"); // Outer region start

            builder.Writeln("Name: <<Name>>");
            builder.Writeln("Country: <<Country>>");
            builder.Writeln("Orders:");

            builder.Writeln("<<TableStart:Orders>>"); // Inner region start
            builder.Writeln("  Order ID: <<OrderID>>, Product: <<Product>>, Qty: <<Quantity>>");
            builder.Writeln("<<TableEnd:Orders>>");   // Inner region end

            builder.Writeln("<<TableEnd:Customers>>"); // Outer region end

            // Execute mail merge with regions using the DataSet.
            doc.MailMerge.ExecuteWithRegions(dataSet);

            // Save the result.
            doc.Save("Output.docx");
        }
    }
}
