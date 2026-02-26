using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MailMergeWithRegionsToPdf
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCX document that contains mail‑merge regions.
            // The Document constructor handles loading; no custom loading code is required.
            Document doc = new Document("TemplateWithRegions.docx");

            // Prepare a DataSet that matches the mail‑merge region names in the template.
            // Each DataTable must have its TableName set to the region name.
            DataSet data = CreateSampleDataSet();

            // Execute the mail merge using the DataSet. This will expand the regions
            // according to the rows in each table.
            doc.MailMerge.ExecuteWithRegions(data);

            // Save the merged document as PDF. The Save method determines the format
            // from the SaveFormat enum; this follows the required lifecycle rule.
            doc.Save("MergedResult.pdf", SaveFormat.Pdf);
        }

        // Creates a sample DataSet with two tables: "Customers" and "Orders".
        // The "Orders" table is related to "Customers" via the "CustomerID" column.
        private static DataSet CreateSampleDataSet()
        {
            // Customers table – corresponds to a mail‑merge region named "Customers".
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("CustomerID", typeof(int));
            customers.Columns.Add("CustomerName", typeof(string));
            customers.Columns.Add("Address", typeof(string));

            customers.Rows.Add(1, "Thomas Hardy", "120 Hanover Sq., London");
            customers.Rows.Add(2, "Paolo Accorti", "Via Monte Bianco 34, Torino");

            // Orders table – corresponds to a nested mail‑merge region named "Orders".
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("CustomerID", typeof(int));
            orders.Columns.Add("ItemName", typeof(string));
            orders.Columns.Add("Quantity", typeof(int));

            orders.Rows.Add(1, "Rugby World Cup Cap", 2);
            orders.Rows.Add(1, "Rugby World Cup Ball", 1);
            orders.Rows.Add(2, "Rugby World Cup Guide", 1);

            // Assemble the DataSet and define the relationship.
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(customers);
            dataSet.Tables.Add(orders);
            dataSet.Relations.Add(customers.Columns["CustomerID"], orders.Columns["CustomerID"]);

            return dataSet;
        }
    }
}
