using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace NestedRegionMailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- Define the outer mail merge region (Customers) ----------
            // Insert the start tag for the outer region.
            builder.InsertField(" MERGEFIELD TableStart:Customers");

            // Add some static text and fields that belong to the outer region.
            builder.Write("Orders for ");
            builder.InsertField(" MERGEFIELD CustomerName");
            builder.Write(":");
            builder.InsertParagraph();

            // ---------- Define a table that will contain the inner region (Orders) ----------
            builder.StartTable();
            // Header row.
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Insert the start tag for the inner region inside the same row.
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD TableStart:Orders");
            builder.InsertField(" MERGEFIELD ItemName");
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD Quantity");

            // End the inner region.
            builder.InsertField(" MERGEFIELD TableEnd:Orders");
            builder.EndTable();

            // End the outer region.
            builder.InsertField(" MERGEFIELD TableEnd:Customers");

            // ---------- Build the hierarchical data source ----------
            DataSet dataSet = CreateCustomersAndOrdersDataSet();

            // Perform the nested mail merge.
            doc.MailMerge.ExecuteWithRegions(dataSet);

            // Save the result.
            doc.Save("NestedMailMergeResult.docx");
        }

        /// <summary>
        /// Creates a DataSet containing two related tables: Customers and Orders.
        /// The Orders table has a many-to-one relationship with Customers on CustomerID.
        /// </summary>
        private static DataSet CreateCustomersAndOrdersDataSet()
        {
            // Customers table.
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("CustomerID", typeof(int));
            customers.Columns.Add("CustomerName", typeof(string));
            customers.Rows.Add(1, "John Doe");
            customers.Rows.Add(2, "Jane Doe");

            // Orders table.
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("CustomerID", typeof(int));
            orders.Columns.Add("ItemName", typeof(string));
            orders.Columns.Add("Quantity", typeof(int));
            orders.Rows.Add(1, "Hawaiian", 2);
            orders.Rows.Add(2, "Pepperoni", 1);
            orders.Rows.Add(2, "Chicago", 1);

            // Assemble the DataSet and define the relationship.
            DataSet ds = new DataSet();
            ds.Tables.Add(customers);
            ds.Tables.Add(orders);
            ds.Relations.Add(customers.Columns["CustomerID"], orders.Columns["CustomerID"]);

            return ds;
        }
    }
}
