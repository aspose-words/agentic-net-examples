using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeToPng
{
    static void Main()
    {
        // Load the DOCX template that contains mail‑merge regions.
        // The constructor overload that takes a file name follows the provided load rule.
        Document doc = new Document("Template.docx");

        // -----------------------------------------------------------------
        // Prepare a DataSet with tables whose names match the region names
        // defined in the template (e.g. TableStart:Customers, TableStart:Orders).
        // -----------------------------------------------------------------
        DataSet dataSet = new DataSet();

        // ----- Customers table (outer region) -----
        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("CustomerID", typeof(int));
        customers.Columns.Add("FullName", typeof(string));
        customers.Columns.Add("Address", typeof(string));

        customers.Rows.Add(1, "Thomas Hardy", "120 Hanover Sq., London");
        customers.Rows.Add(2, "Paolo Accorti", "Via Monte Bianco 34, Torino");

        dataSet.Tables.Add(customers);

        // ----- Orders table (nested region) -----
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("CustomerID", typeof(int));
        orders.Columns.Add("ItemName", typeof(string));
        orders.Columns.Add("Quantity", typeof(int));

        orders.Rows.Add(1, "Rugby World Cup Cap", 2);
        orders.Rows.Add(1, "Rugby World Cup Ball", 1);
        orders.Rows.Add(2, "Rugby World Cup Guide", 1);

        dataSet.Tables.Add(orders);

        // Define the relationship between the two tables so that the nested region
        // (Orders) can be linked to its parent (Customers) by CustomerID.
        dataSet.Relations.Add(customers.Columns["CustomerID"], orders.Columns["CustomerID"]);

        // -----------------------------------------------------------------
        // Execute the mail merge using the DataSet.
        // The ExecuteWithRegions method is the rule‑based way to merge data
        // into repeatable regions.
        // -----------------------------------------------------------------
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // -----------------------------------------------------------------
        // Save the merged document as PNG.
        // Use ImageSaveOptions with SaveFormat.Png and the Save(string, SaveOptions) overload.
        // -----------------------------------------------------------------
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        // Optional: render only the first page (remove the line to render all pages).
        // pngOptions.PageSet = new PageSet(0);

        doc.Save("MergedResult.png", pngOptions);
    }
}
