using System;
using System.Data;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the outer mail merge region for orders.
        builder.InsertField(" MERGEFIELD TableStart:Orders");
        builder.Writeln("Order ID: ");
        builder.InsertField(" MERGEFIELD OrderID");
        builder.Writeln();
        builder.Writeln("Customer: ");
        builder.InsertField(" MERGEFIELD CustomerName");
        builder.Writeln();

        // Define the inner mail merge region for products belonging to the current order.
        builder.InsertField(" MERGEFIELD TableStart:Products");
        builder.Writeln("\tProduct: ");
        builder.InsertField(" MERGEFIELD ProductName");
        builder.Write(", Qty: ");
        builder.InsertField(" MERGEFIELD Quantity");
        builder.Writeln();
        builder.InsertField(" MERGEFIELD TableEnd:Products");

        // Close the outer region.
        builder.InsertField(" MERGEFIELD TableEnd:Orders");

        // Build the data set with two related tables.
        DataSet dataSet = CreateDataSet();

        // Perform mail merge with regions.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the result.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OrdersReport.docx");
        doc.Save(outputPath);
    }

    // Creates a DataSet containing Orders and Products tables with a relation on OrderID.
    private static DataSet CreateDataSet()
    {
        // Orders table.
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderID", typeof(int));
        orders.Columns.Add("CustomerName", typeof(string));
        orders.Rows.Add(1, "Alice Johnson");
        orders.Rows.Add(2, "Bob Smith");

        // Products table (order items).
        DataTable products = new DataTable("Products");
        products.Columns.Add("OrderID", typeof(int));
        products.Columns.Add("ProductName", typeof(string));
        products.Columns.Add("Quantity", typeof(int));
        products.Rows.Add(1, "Laptop", 1);
        products.Rows.Add(1, "Mouse", 2);
        products.Rows.Add(2, "Desk", 1);
        products.Rows.Add(2, "Chair", 4);
        products.Rows.Add(2, "Lamp", 2);

        // Create the DataSet and add tables.
        DataSet ds = new DataSet();
        ds.Tables.Add(orders);
        ds.Tables.Add(products);

        // Define the relation between Orders and Products on OrderID.
        ds.Relations.Add("Order_Products",
            orders.Columns["OrderID"],
            products.Columns["OrderID"]);

        return ds;
    }
}
