using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class DotmContextualAccessExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple DOTM template that accesses object members using the <<[Object.Member]>> syntax.
        // The template will display a customer's name and list all of their order product names.
        builder.Writeln("Customer Name: <<[Customer.Name]>>");
        builder.Writeln("Orders:");
        builder.Writeln("<<foreach [in Customer.Orders]>>- <<[Product]>>");
        builder.Writeln("<</foreach>>");

        // Prepare a DataSet that mimics the object hierarchy used in the template.
        // The DataSet contains a "Customer" table with a "Name" column,
        // and an "Orders" table with a "Product" column linked to the customer via "CustomerId".
        DataSet dataSet = new DataSet();

        // Customer table.
        DataTable customerTable = new DataTable("Customer");
        customerTable.Columns.Add("Id", typeof(int));
        customerTable.Columns.Add("Name", typeof(string));
        customerTable.Rows.Add(1, "John Doe");
        dataSet.Tables.Add(customerTable);

        // Orders table.
        DataTable ordersTable = new DataTable("Orders");
        ordersTable.Columns.Add("CustomerId", typeof(int));
        ordersTable.Columns.Add("Product", typeof(string));
        ordersTable.Rows.Add(1, "Laptop");
        ordersTable.Rows.Add(1, "Smartphone");
        dataSet.Tables.Add(ordersTable);

        // Establish a relation so the ReportingEngine can navigate Customer.Orders.
        dataSet.Relations.Add("Customer_Orders",
            customerTable.Columns["Id"],
            ordersTable.Columns["CustomerId"]);

        // Configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            // Allow the engine to continue even if a member is missing.
            Options = ReportBuildOptions.AllowMissingMembers,
            // Text to display for any missing member.
            MissingMemberMessage = "N/A"
        };

        // Build the report using the template document and the prepared DataSet.
        engine.BuildReport(doc, dataSet, string.Empty);

        // Save the resulting document.
        doc.Save("Output.dotm");
    }
}
