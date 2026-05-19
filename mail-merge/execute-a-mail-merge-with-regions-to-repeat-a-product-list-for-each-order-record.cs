using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeWithRegionsExample
{
    class Program
    {
        static void Main()
        {
            // Create a blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- Define mail merge regions ----------
            // Outer region: Orders
            builder.InsertField(" MERGEFIELD TableStart:Orders");
            builder.Writeln("Order ID: ");
            builder.InsertField(" MERGEFIELD OrderID");
            builder.Writeln("Customer: ");
            builder.InsertField(" MERGEFIELD CustomerName");
            builder.Writeln("Products:");

            // Inner region: Products (nested inside Orders)
            builder.InsertField(" MERGEFIELD TableStart:Products");
            builder.InsertField(" MERGEFIELD ProductName");
            builder.Write("\tQty: ");
            builder.InsertField(" MERGEFIELD Quantity");
            builder.InsertParagraph();

            // End inner region
            builder.InsertField(" MERGEFIELD TableEnd:Products");
            // End outer region
            builder.InsertField(" MERGEFIELD TableEnd:Orders");

            // ---------- Build the data source ----------
            // Orders table
            DataTable ordersTable = new DataTable("Orders");
            ordersTable.Columns.Add("OrderID", typeof(int));
            ordersTable.Columns.Add("CustomerName", typeof(string));
            ordersTable.Rows.Add(1, "Alice");
            ordersTable.Rows.Add(2, "Bob");

            // Products table (many-to-one relationship with Orders)
            DataTable productsTable = new DataTable("Products");
            productsTable.Columns.Add("OrderID", typeof(int)); // foreign key
            productsTable.Columns.Add("ProductName", typeof(string));
            productsTable.Columns.Add("Quantity", typeof(int));
            productsTable.Rows.Add(1, "Apple", 5);
            productsTable.Rows.Add(1, "Banana", 3);
            productsTable.Rows.Add(2, "Orange", 2);
            productsTable.Rows.Add(2, "Grapes", 4);
            productsTable.Rows.Add(2, "Mango", 1);

            // Create a DataSet and add the tables.
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(ordersTable);
            dataSet.Tables.Add(productsTable);

            // Define the relationship between Orders and Products on OrderID.
            dataSet.Relations.Add(
                ordersTable.Columns["OrderID"],
                productsTable.Columns["OrderID"]);

            // ---------- Execute mail merge with regions ----------
            doc.MailMerge.ExecuteWithRegions(dataSet);

            // ---------- Save the result ----------
            string outputPath = Path.Combine(Environment.CurrentDirectory, "MailMergeWithRegionsOutput.docx");
            doc.Save(outputPath);
        }
    }
}
