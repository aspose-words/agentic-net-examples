using System;
using System.Data;
using Aspose.Words;

namespace MailMergeWithRegionsExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains mail merge regions.
            // The template must have TableStart: and TableEnd: fields that define the regions.
            Document doc = new Document("TemplateWithRegions.docx");

            // Create a DataSet that holds the data for each region.
            // Each DataTable name must match the region name used in the template.
            DataSet data = CreateDataSet();

            // Perform the mail merge with regions.
            // This will repeat the content inside each region for every row in the corresponding table.
            doc.MailMerge.ExecuteWithRegions(data);

            // Save the merged document.
            doc.Save("MergedResult.docx");
        }

        // Helper method to build a DataSet with two related tables: Customers and Orders.
        private static DataSet CreateDataSet()
        {
            // Customers table – will be used for the outer mail merge region.
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("CustomerID", typeof(int));
            customers.Columns.Add("CustomerName", typeof(string));
            customers.Columns.Add("Address", typeof(string));
            customers.Rows.Add(1, "John Doe", "120 Hanover Sq., London");
            customers.Rows.Add(2, "Jane Smith", "5th Avenue, New York");

            // Orders table – will be used for the nested mail merge region.
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("CustomerID", typeof(int));
            orders.Columns.Add("Item", typeof(string));
            orders.Columns.Add("Quantity", typeof(int));
            orders.Rows.Add(1, "Rugby Ball", 2);
            orders.Rows.Add(1, "Cap", 1);
            orders.Rows.Add(2, "Jersey", 3);

            // Create the DataSet and add the tables.
            DataSet ds = new DataSet();
            ds.Tables.Add(customers);
            ds.Tables.Add(orders);

            // Define the relationship between Customers and Orders on CustomerID.
            ds.Relations.Add(customers.Columns["CustomerID"], orders.Columns["CustomerID"]);

            return ds;
        }
    }
}
