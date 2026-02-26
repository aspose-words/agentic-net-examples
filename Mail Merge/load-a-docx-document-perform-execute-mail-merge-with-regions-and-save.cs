using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains mail‑merge regions.
        Document doc = new Document("Template.docx");

        // ------------------------------------------------------------
        // Prepare a DataSet with tables whose names match the region names
        // defined in the template (e.g. TableStart:Customers, TableStart:Orders).
        // ------------------------------------------------------------
        DataSet dataSet = new DataSet();

        // First region: Customers
        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("FullName");
        customers.Columns.Add("Address");
        customers.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");
        customers.Rows.Add("Paolo Accorti", "Via Monte Bianco 34, Torino");
        dataSet.Tables.Add(customers);

        // Second region: Orders (related to Customers by FullName)
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("FullName");   // foreign key to Customers.FullName
        orders.Columns.Add("Item");
        orders.Columns.Add("Quantity");
        orders.Rows.Add("Thomas Hardy", "Rugby World Cup Cap", "2");
        orders.Rows.Add("Thomas Hardy", "Rugby World Cup Ball", "1");
        orders.Rows.Add("Paolo Accorti", "Rugby World Cup Guide", "1");
        dataSet.Tables.Add(orders);

        // Define the relationship so that nested regions work correctly.
        dataSet.Relations.Add("CustOrders",
            customers.Columns["FullName"],
            orders.Columns["FullName"]);

        // ------------------------------------------------------------
        // Execute mail merge with regions using the prepared DataSet.
        // ------------------------------------------------------------
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // ------------------------------------------------------------
        // Save the merged document as a JPEG image.
        // ------------------------------------------------------------
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        // Render only the first page (index 0). Remove this line to render all pages.
        jpegOptions.PageSet = new PageSet(0);
        doc.Save("MergedResult.jpg", jpegOptions);
    }
}
