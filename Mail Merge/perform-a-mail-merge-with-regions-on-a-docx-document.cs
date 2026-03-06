using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeWithRegionsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -------------------------------------------------
            // Define the first mail merge region: Customers.
            // -------------------------------------------------
            builder.InsertField(" MERGEFIELD TableStart:Customers");
            builder.Writeln("Customer Name: ");
            builder.InsertField(" MERGEFIELD FullName");
            builder.Writeln();
            builder.Writeln("Address: ");
            builder.InsertField(" MERGEFIELD Address");
            builder.Writeln();

            // Insert a nested region for Orders belonging to each customer.
            builder.InsertField(" MERGEFIELD TableStart:Orders");
            builder.Writeln("\tOrder Item: ");
            builder.InsertField(" MERGEFIELD ItemName");
            builder.Writeln("\tQuantity: ");
            builder.InsertField(" MERGEFIELD Quantity");
            builder.Writeln();
            builder.InsertField(" MERGEFIELD TableEnd:Orders");

            // End the Customers region.
            builder.InsertField(" MERGEFIELD TableEnd:Customers");

            // -------------------------------------------------
            // Prepare the data source: a DataSet with two related tables.
            // -------------------------------------------------
            DataSet data = CreateDataSet();

            // Perform the mail merge with regions.
            doc.MailMerge.ExecuteWithRegions(data);

            // Save the merged document.
            doc.Save("MailMergeWithRegionsResult.docx");
        }

        /// <summary>
        /// Creates a DataSet containing a Customers table and an Orders table.
        /// The Orders table has a foreign key relationship to Customers on CustomerID.
        /// </summary>
        private static DataSet CreateDataSet()
        {
            // Customers table.
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("CustomerID", typeof(int));
            customers.Columns.Add("FullName", typeof(string));
            customers.Columns.Add("Address", typeof(string));
            customers.Rows.Add(1, "Thomas Hardy", "120 Hanover Sq., London");
            customers.Rows.Add(2, "Paolo Accorti", "Via Monte Bianco 34, Torino");

            // Orders table.
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("CustomerID", typeof(int));
            orders.Columns.Add("ItemName", typeof(string));
            orders.Columns.Add("Quantity", typeof(int));
            orders.Rows.Add(1, "Rugby World Cup Cap", 2);
            orders.Rows.Add(1, "Rugby World Cup Ball", 1);
            orders.Rows.Add(2, "Rugby World Cup Guide", 1);
            orders.Rows.Add(2, "Rugby World Cup Shirt", 3);

            // Create the DataSet and add the tables.
            DataSet ds = new DataSet();
            ds.Tables.Add(customers);
            ds.Tables.Add(orders);

            // Define the relationship between Customers and Orders.
            ds.Relations.Add(customers.Columns["CustomerID"], orders.Columns["CustomerID"]);

            return ds;
        }
    }
}
