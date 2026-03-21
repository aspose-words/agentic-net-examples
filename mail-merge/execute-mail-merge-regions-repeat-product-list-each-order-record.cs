using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    class Program
    {
        static void Main()
        {
            MailMergeWithRegionsExample.Run();
        }
    }

    public static class MailMergeWithRegionsExample
    {
        public static void Run()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- Begin outer mail merge region "Orders" ----------
            builder.InsertField(" MERGEFIELD TableStart:Orders");
            builder.Write("Order ID: ");
            builder.InsertField(" MERGEFIELD OrderID");
            builder.Writeln();

            // ---------- Begin inner mail merge region "Products" ----------
            builder.InsertField(" MERGEFIELD TableStart:Products");

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

            DataSet data = CreateDataSet();

            doc.MailMerge.ExecuteWithRegions(data);

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

            // Assemble the DataSet.
            DataSet ds = new DataSet();
            ds.Tables.Add(orders);
            ds.Tables.Add(products);

            // Define the relation between Orders and Products on OrderID.
            DataRelation relation = new DataRelation(
                "Orders_Products",
                ds.Tables["Orders"]!.Columns["OrderID"]!,
                ds.Tables["Products"]!.Columns["OrderID"]!
            );

            ds.Relations.Add(relation);

            return ds;
        }
    }
}
