using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

public class MailMergeWithRegionsExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a simple field for OrderID.
        builder.Writeln("Order ID: <<OrderID>>");
        builder.Writeln();

        // Start the mail merge region for the product list.
        builder.Writeln("<<ProductStart>>");
        // Create a table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Product Name");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Insert placeholders for product fields.
        builder.InsertCell();
        builder.Write("<<ProductName>>");
        builder.InsertCell();
        builder.Write("<<Quantity>>");
        builder.EndRow();
        builder.EndTable();
        // End the mail merge region.
        builder.Writeln("<<ProductEnd>>");

        // Prepare data for mail merge.
        DataSet dataSet = new DataSet();

        // Orders table (master).
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderID", typeof(int));
        orders.Rows.Add(1001);
        orders.Rows.Add(1002);
        dataSet.Tables.Add(orders);

        // Products table (detail).
        DataTable products = new DataTable("Product");
        products.Columns.Add("OrderID", typeof(int));
        products.Columns.Add("ProductName", typeof(string));
        products.Columns.Add("Quantity", typeof(int));
        // Products for first order.
        products.Rows.Add(1001, "Apple", 5);
        products.Rows.Add(1001, "Banana", 3);
        // Products for second order.
        products.Rows.Add(1002, "Orange", 7);
        products.Rows.Add(1002, "Grapes", 2);
        dataSet.Tables.Add(products);

        // Define relation between Orders and Product tables.
        dataSet.Relations.Add("Order_Product",
            orders.Columns["OrderID"],
            products.Columns["OrderID"]);

        // Execute mail merge with regions.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the result to a file.
        string outputPath = "MailMergeWithRegionsOutput.docx";
        doc.Save(outputPath);

        // Optionally, indicate completion (no interactive prompts).
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
