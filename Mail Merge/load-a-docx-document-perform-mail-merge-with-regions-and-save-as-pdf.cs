using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class MailMergeWithRegionsToPdf
{
    static void Main()
    {
        // Load the DOCX template that contains mail‑merge regions.
        Document doc = new Document("Template.docx");

        // ------------------------------------------------------------
        // Prepare a DataSet with tables that match the region names.
        // ------------------------------------------------------------
        DataSet dataSet = new DataSet();

        // First region: Customers
        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("CustomerName");
        customers.Columns.Add("Address");
        customers.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");
        customers.Rows.Add("Paolo Accorti", "Via Monte Bianco 34, Torino");
        dataSet.Tables.Add(customers);

        // Second (nested) region: Orders
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("CustomerName"); // foreign key to link with Customers
        orders.Columns.Add("Item");
        orders.Columns.Add("Quantity");
        orders.Rows.Add("Thomas Hardy", "Rugby World Cup Cap", "2");
        orders.Rows.Add("Thomas Hardy", "Rugby World Cup Ball", "1");
        orders.Rows.Add("Paolo Accorti", "Rugby World Cup Guide", "1");
        dataSet.Tables.Add(orders);

        // Define the relationship so that Aspose.Words can repeat the inner region per parent row.
        dataSet.Relations.Add(customers.Columns["CustomerName"], orders.Columns["CustomerName"]);

        // ------------------------------------------------------------
        // Execute mail merge with regions using the prepared DataSet.
        // ------------------------------------------------------------
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // ------------------------------------------------------------
        // Save the merged document as PDF.
        // ------------------------------------------------------------
        doc.Save("Result.pdf", SaveFormat.Pdf);
    }
}
