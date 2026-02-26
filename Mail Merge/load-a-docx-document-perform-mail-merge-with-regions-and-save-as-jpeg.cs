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
        // Prepare a DataSet whose tables have the same names as the
        // mail‑merge regions defined in the template (e.g. TableStart:Customers,
        // TableStart:Orders).  The fields inside each region must match the
        // column names of the corresponding table.
        // ------------------------------------------------------------
        DataSet data = new DataSet();

        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("FullName");
        customers.Columns.Add("Address");
        customers.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");
        customers.Rows.Add("Paolo Accorti", "Via Monte Bianco 34, Torino");
        data.Tables.Add(customers);

        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("Item");
        orders.Columns.Add("Quantity");
        orders.Rows.Add("Rugby World Cup Cap", "2");
        orders.Rows.Add("Rugby World Cup Ball", "1");
        data.Tables.Add(orders);

        // Execute the mail merge with regions.  The document will expand the
        // regions to accommodate all rows in each table.
        doc.MailMerge.ExecuteWithRegions(data);

        // ------------------------------------------------------------
        // Save the merged document as a JPEG image.
        // ImageSaveOptions allows us to specify the output format and
        // which page(s) to render.  Here we render only the first page.
        // ------------------------------------------------------------
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        jpegOptions.PageSet = new PageSet(0); // zero‑based index of the first page
        doc.Save("MergedResult.jpg", jpegOptions);
    }
}
