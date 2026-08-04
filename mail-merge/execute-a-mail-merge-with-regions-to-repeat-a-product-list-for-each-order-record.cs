using System;
using System.Data;
using Aspose.Words;

public class MailMergeWithRegionsExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Define mail merge regions ----------
        // Outer region: Orders
        builder.InsertField(" MERGEFIELD TableStart:Orders");
        builder.Writeln("Order ID: ");
        builder.InsertField(" MERGEFIELD OrderID");
        builder.Writeln();
        builder.Writeln("Customer: ");
        builder.InsertField(" MERGEFIELD CustomerName");
        builder.Writeln();
        builder.Writeln("Products:");
        // Inner region: Products (related to the current order)
        builder.InsertField(" MERGEFIELD TableStart:Products");
        builder.Write("\tProduct: ");
        builder.InsertField(" MERGEFIELD ProductName");
        builder.Write(", Qty: ");
        builder.InsertField(" MERGEFIELD Quantity");
        builder.Writeln();
        builder.InsertField(" MERGEFIELD TableEnd:Products");
        // End of Orders region
        builder.InsertField(" MERGEFIELD TableEnd:Orders");

        // ---------- Prepare data ----------
        DataSet dataSet = CreateDataSet();

        // ---------- Perform mail merge with regions ----------
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // ---------- Save the result ----------
        doc.Save("MailMergeWithRegionsOutput.docx");
    }

    // Creates a DataSet containing Orders and Products tables with a relation.
    private static DataSet CreateDataSet()
    {
        // Orders table (master)
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderID", typeof(int));
        orders.Columns.Add("CustomerName", typeof(string));
        orders.Rows.Add(1, "John Doe");
        orders.Rows.Add(2, "Jane Smith");

        // Products table (detail)
        DataTable products = new DataTable("Products");
        products.Columns.Add("OrderID", typeof(int));
        products.Columns.Add("ProductName", typeof(string));
        products.Columns.Add("Quantity", typeof(int));
        // Products for Order 1
        products.Rows.Add(1, "Laptop", 1);
        products.Rows.Add(1, "Mouse", 2);
        // Products for Order 2
        products.Rows.Add(2, "Desk", 1);
        products.Rows.Add(2, "Chair", 4);
        products.Rows.Add(2, "Lamp", 2);

        // Create DataSet and add tables.
        DataSet ds = new DataSet();
        ds.Tables.Add(orders);
        ds.Tables.Add(products);

        // Define relation between Orders and Products on OrderID.
        ds.Relations.Add("Order_Products",
            orders.Columns["OrderID"],
            products.Columns["OrderID"]);

        return ds;
    }
}
