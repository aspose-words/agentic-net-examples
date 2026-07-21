using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class MailMergeWithRegionsExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Define the outer mail merge region "Orders" -----
        // This region will be repeated for each order record.
        builder.InsertField(" MERGEFIELD TableStart:Orders");

        // Fields inside the "Orders" region.
        builder.Write("Order ID: ");
        builder.InsertField(" MERGEFIELD OrderID");
        builder.Write("\nCustomer: ");
        builder.InsertField(" MERGEFIELD CustomerName");
        builder.Writeln("\nProducts:");

        // ----- Define the inner mail merge region "Products" -----
        // This region will be repeated for each product belonging to the current order.
        // We'll place it inside a table for nicer formatting.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Product Name");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Begin the "Products" region.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD TableStart:Products");
        builder.InsertField(" MERGEFIELD ProductName");
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Quantity");
        // End the "Products" region.
        builder.InsertField(" MERGEFIELD TableEnd:Products");
        builder.EndTable();

        // End the outer "Orders" region.
        builder.InsertField(" MERGEFIELD TableEnd:Orders");

        // ----- Prepare the data source -----
        // Create a DataSet containing two related tables: Orders and Products.
        DataSet data = CreateDataSet();

        // Execute the mail merge with regions using the DataSet.
        doc.MailMerge.ExecuteWithRegions(data);

        // Save the resulting document.
        doc.Save("MailMergeWithRegionsOutput.docx");
    }

    // Generates a DataSet with sample orders and their corresponding products.
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
        products.Rows.Add(1, "Apple", 5);
        products.Rows.Add(1, "Banana", 3);
        products.Rows.Add(2, "Orange", 7);
        products.Rows.Add(2, "Grapes", 2);
        products.Rows.Add(2, "Mango", 4);

        // Create the DataSet and add the tables.
        DataSet ds = new DataSet();
        ds.Tables.Add(orders);
        ds.Tables.Add(products);

        // Define a relation between Orders and Products on OrderID.
        ds.Relations.Add("Order_Products",
            orders.Columns["OrderID"],
            products.Columns["OrderID"]);

        return ds;
    }
}
