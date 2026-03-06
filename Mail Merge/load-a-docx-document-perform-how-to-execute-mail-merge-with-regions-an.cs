using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX that contains mail‑merge regions.
        // The document must have fields like TableStart:Customers / TableEnd:Customers, etc.
        Document doc = new Document("TemplateWithRegions.docx");

        // Prepare a DataSet whose table names match the region names in the document.
        DataSet dataSet = CreateDataSet();

        // Execute mail merge with regions using the prepared DataSet.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the merged document as a PNG image.
        // Each page of the document will be rendered to a separate PNG file.
        // Here we render only the first page for simplicity.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new PageSet(0) // zero‑based index of the page to render
        };
        doc.Save("MergedResult.png", options);
    }

    // Creates a DataSet containing two related tables: Customers and Orders.
    // The table names must correspond to the mail‑merge region names in the template.
    private static DataSet CreateDataSet()
    {
        // Customers table (region name: Customers)
        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("CustomerID", typeof(int));
        customers.Columns.Add("CustomerName", typeof(string));
        customers.Columns.Add("Address", typeof(string));

        customers.Rows.Add(1, "Thomas Hardy", "120 Hanover Sq., London");
        customers.Rows.Add(2, "Paolo Accorti", "Via Monte Bianco 34, Torino");

        // Orders table (region name: Orders)
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("CustomerID", typeof(int));
        orders.Columns.Add("ItemName", typeof(string));
        orders.Columns.Add("Quantity", typeof(int));

        orders.Rows.Add(1, "Rugby World Cup Cap", 2);
        orders.Rows.Add(1, "Rugby World Cup Ball", 1);
        orders.Rows.Add(2, "Rugby World Cup Guide", 1);

        // Assemble the DataSet and define the relationship.
        DataSet ds = new DataSet();
        ds.Tables.Add(customers);
        ds.Tables.Add(orders);
        ds.Relations.Add(customers.Columns["CustomerID"], orders.Columns["CustomerID"]);

        return ds;
    }
}
