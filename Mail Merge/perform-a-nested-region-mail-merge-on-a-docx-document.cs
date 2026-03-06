using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace NestedMailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();                     // create
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- Define outer mail merge region (Customers) ----------
            // TableStart:Customers marks the beginning of the region.
            builder.InsertField(" MERGEFIELD TableStart:Customers");

            // Fields inside the outer region.
            builder.Write("Orders for ");
            builder.InsertField(" MERGEFIELD CustomerName");
            builder.Write(":");
            builder.InsertParagraph();

            // Create a table to hold the inner region (Orders).
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // ---------- Define inner mail merge region (Orders) ----------
            // The inner region must start and end on the same table row.
            builder.InsertCell(); // move to next row, first cell
            builder.InsertField(" MERGEFIELD TableStart:Orders");
            builder.InsertField(" MERGEFIELD ItemName");
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD Quantity");
            builder.InsertField(" MERGEFIELD TableEnd:Orders");
            builder.EndTable();

            // End the outer region.
            builder.InsertField(" MERGEFIELD TableEnd:Customers");

            // ---------- Prepare hierarchical data (Customers -> Orders) ----------
            DataSet dataSet = CreateCustomersOrdersDataSet();

            // Perform the nested mail merge using the DataSet.
            doc.MailMerge.ExecuteWithRegions(dataSet);          // execute with regions

            // Save the merged document.
            doc.Save("NestedMailMerge.docx");                  // save
        }

        /// <summary>
        /// Creates a DataSet containing two related tables:
        /// "Customers" (parent) and "Orders" (child) linked by CustomerID.
        /// </summary>
        private static DataSet CreateCustomersOrdersDataSet()
        {
            // ----- Customers table -----
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("CustomerID", typeof(int));
            customers.Columns.Add("CustomerName", typeof(string));
            customers.Rows.Add(1, "John Doe");
            customers.Rows.Add(2, "Jane Doe");

            // ----- Orders table -----
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("CustomerID", typeof(int));
            orders.Columns.Add("ItemName", typeof(string));
            orders.Columns.Add("Quantity", typeof(int));
            orders.Rows.Add(1, "Hawaiian", 2);
            orders.Rows.Add(2, "Pepperoni", 1);
            orders.Rows.Add(2, "Chicago", 1);

            // ----- Build DataSet with relationship -----
            DataSet ds = new DataSet();
            ds.Tables.Add(customers);
            ds.Tables.Add(orders);
            ds.Relations.Add(customers.Columns["CustomerID"], orders.Columns["CustomerID"]);

            return ds;
        }
    }
}
