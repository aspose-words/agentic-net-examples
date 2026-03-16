using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    class Program
    {
        // Entry point required by the C# compiler.
        static void Main()
        {
            MailMergeWithRegionsExample.Run();
        }
    }

    public static class MailMergeWithRegionsExample
    {
        public static void Run()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- Begin outer mail merge region "Orders" ----------
            // This region will be repeated for each row in the Orders table.
            builder.InsertField(" MERGEFIELD TableStart:Orders");
            builder.Write("Order ID: ");
            builder.InsertField(" MERGEFIELD OrderID");
            builder.Writeln();

            // ---------- Begin inner mail merge region "Products" ----------
            // This region will be repeated for each product belonging to the current order.
            builder.InsertField(" MERGEFIELD TableStart:Products");

            // Create a simple table to display product name and quantity.
            builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Data row – will be duplicated for each product record.
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD ProductName");
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD Quantity");
            builder.EndRow();

            builder.EndTable();

            // End inner region.
            builder.InsertField(" MERGEFIELD TableEnd:Products");
            // End outer region.
            builder.InsertField(" MERGEFIELD TableEnd:Orders");

            // Build the data source: a DataSet containing two related tables.
            DataSet data = CreateDataSet();

            // Execute the mail merge with regions. The document will expand the regions
            // according to the rows in the DataSet.
            doc.MailMerge.ExecuteWithRegions(data);

            // Save the merged document.
            doc.Save("OrdersWithProducts.docx");
        }

        private static DataSet CreateDataSet()
        {
            // ----- Orders table (master) -----
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Rows.Add(1001);
            orders.Rows.Add(1002);

            // ----- Products table (detail) -----
            DataTable products = new DataTable("Products");
            products.Columns.Add("OrderID", typeof(int));
            products.Columns.Add("ProductName", typeof(string));
            products.Columns.Add("Quantity", typeof(int));

            // Products for Order 1001
            products.Rows.Add(1001, "Apple", 3);
            products.Rows.Add(1001, "Banana", 5);

            // Products for Order 1002
            products.Rows.Add(1002, "Orange", 2);
            products.Rows.Add(1002, "Grapes", 1);
            products.Rows.Add(1002, "Mango", 4);

            // Define the relation between Orders and Products on OrderID.
            // The null‑forgiving operator (!) tells the compiler the columns exist.
            DataRelation relation = new DataRelation(
                "Orders_Products",
                orders.Columns["OrderID"]!,
                products.Columns["OrderID"]!
            );

            // Assemble the DataSet.
            DataSet ds = new DataSet();
            ds.Tables.Add(orders);
            ds.Tables.Add(products);
            ds.Relations.Add(relation);

            return ds;
        }
    }
}
