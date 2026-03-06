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
            // Load the DOCX template that contains mail merge regions.
            // The template must have pairs of MERGEFIELDs with TableStart:RegionName and TableEnd:RegionName.
            Document doc = new Document("TemplateWithRegions.docx");

            // Build a DataSet that defines the data for each region.
            // Each DataTable name must match the region name used in the template.
            DataSet dataSet = CreateDataSet();

            // Perform the mail merge with regions.
            // This will repeat the content inside each region for every row in the corresponding table.
            doc.MailMerge.ExecuteWithRegions(dataSet);

            // Save the merged document.
            doc.Save("MergedResult.docx");
        }

        /// <summary>
        /// Creates a DataSet containing two tables: "Customers" and "Orders".
        /// The "Orders" table has a foreign key relationship to "Customers" on CustomerID.
        /// </summary>
        private static DataSet CreateDataSet()
        {
            // Customers table – will be used for the outer mail merge region.
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("CustomerID", typeof(int));
            customers.Columns.Add("CustomerName", typeof(string));
            customers.Rows.Add(1, "John Doe");
            customers.Rows.Add(2, "Jane Smith");

            // Orders table – will be used for the nested mail merge region.
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("CustomerID", typeof(int));
            orders.Columns.Add("ItemName", typeof(string));
            orders.Columns.Add("Quantity", typeof(int));
            orders.Rows.Add(1, "Laptop", 1);
            orders.Rows.Add(1, "Mouse", 2);
            orders.Rows.Add(2, "Keyboard", 1);
            orders.Rows.Add(2, "Monitor", 2);

            // Create the DataSet and add the tables.
            DataSet ds = new DataSet();
            ds.Tables.Add(customers);
            ds.Tables.Add(orders);

            // Define the relationship so that Aspose.Words can perform nested mail merge.
            ds.Relations.Add("CustomerOrders",
                customers.Columns["CustomerID"],
                orders.Columns["CustomerID"]);

            return ds;
        }
    }
}
