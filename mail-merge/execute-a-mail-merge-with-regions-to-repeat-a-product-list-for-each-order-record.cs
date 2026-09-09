using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Define the outer mail merge region "Orders" -----
        builder.InsertField(" MERGEFIELD TableStart:Orders");
        builder.Writeln("Order ID: ");
        builder.InsertField(" MERGEFIELD OrderID");
        builder.Writeln();
        builder.Writeln("Customer: ");
        builder.InsertField(" MERGEFIELD CustomerName");
        builder.Writeln();

        // ----- Define the inner mail merge region "Products" -----
        builder.InsertField(" MERGEFIELD TableStart:Products");
        builder.Writeln("\tProduct: ");
        builder.InsertField(" MERGEFIELD ProductName");
        builder.Writeln("\tQuantity: ");
        builder.InsertField(" MERGEFIELD Quantity");
        builder.Writeln();
        builder.InsertField(" MERGEFIELD TableEnd:Products");

        // End the outer region.
        builder.InsertField(" MERGEFIELD TableEnd:Orders");

        // Build the data set with two related tables.
        DataSet dataSet = CreateDataSet();

        // Perform the mail merge with regions.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the result.
        doc.Save("MailMergeWithRegions.docx");
    }

    // Creates a DataSet containing "Orders" and "Products" tables with a one‑to‑many relationship.
    private static DataSet CreateDataSet()
    {
        // Orders table.
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderID", typeof(int));
        orders.Columns.Add("CustomerName", typeof(string));
        orders.Rows.Add(1, "John Doe");
        orders.Rows.Add(2, "Jane Smith");

        // Products table.
        DataTable products = new DataTable("Products");
        products.Columns.Add("OrderID", typeof(int));
        products.Columns.Add("ProductName", typeof(string));
        products.Columns.Add("Quantity", typeof(int));
        products.Rows.Add(1, "Laptop", 1);
        products.Rows.Add(1, "Mouse", 2);
        products.Rows.Add(2, "Keyboard", 1);
        products.Rows.Add(2, "Monitor", 2);
        products.Rows.Add(2, "USB‑Cable", 5);

        // Create the DataSet and add the tables.
        DataSet ds = new DataSet();
        ds.Tables.Add(orders);
        ds.Tables.Add(products);

        // Define the relationship between Orders and Products on OrderID.
        ds.Relations.Add("Order_Products",
            orders.Columns["OrderID"],
            products.Columns["OrderID"]);

        return ds;
    }
}
