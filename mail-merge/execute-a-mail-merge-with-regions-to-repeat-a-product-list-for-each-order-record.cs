using System;
using System.Data;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Define mail‑merge regions ----------
        // Outer region: Orders
        builder.InsertField(" MERGEFIELD TableStart:Orders");
        builder.Write("Order ID: ");
        builder.InsertField(" MERGEFIELD OrderID");
        builder.Write(", Customer: ");
        builder.InsertField(" MERGEFIELD CustomerName");
        builder.Writeln();

        // Inner region: Products (repeated for each order)
        builder.InsertField(" MERGEFIELD TableStart:Products");
        builder.Write("\tProduct: ");
        builder.InsertField(" MERGEFIELD ProductName");
        builder.Write(", Qty: ");
        builder.InsertField(" MERGEFIELD Quantity");
        builder.InsertParagraph();
        builder.InsertField(" MERGEFIELD TableEnd:Products");

        // Close the outer region.
        builder.InsertField(" MERGEFIELD TableEnd:Orders");

        // ---------- Prepare data ----------
        // Orders table (master)
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderID", typeof(int));
        orders.Columns.Add("CustomerName", typeof(string));
        orders.Rows.Add(1, "Alice");
        orders.Rows.Add(2, "Bob");

        // Products table (detail)
        DataTable products = new DataTable("Products");
        products.Columns.Add("OrderID", typeof(int));
        products.Columns.Add("ProductName", typeof(string));
        products.Columns.Add("Quantity", typeof(int));
        products.Rows.Add(1, "Apple", 5);
        products.Rows.Add(1, "Banana", 3);
        products.Rows.Add(2, "Orange", 2);
        products.Rows.Add(2, "Grapes", 4);

        // Create a DataSet that contains both tables and a relation between them.
        DataSet dataSet = new DataSet();
        dataSet.Tables.Add(orders);
        dataSet.Tables.Add(products);
        dataSet.Relations.Add("Order_Products",
            orders.Columns["OrderID"],
            products.Columns["OrderID"]);

        // ---------- Execute mail merge with regions ----------
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the result.
        doc.Save("MailMergeWithRegions.docx");
    }
}
