using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Template.docx");

        // Create a DataTable that represents a mail‑merge region named "Customers".
        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("FullName");
        customers.Columns.Add("Address");
        customers.Rows.Add("John Doe", "123 Main St, Anytown");
        customers.Rows.Add("Jane Smith", "456 Oak Ave, Othertown");

        // Execute mail merge with regions. The table name must match the region name
        // defined in the document (e.g. <<TableStart:Customers>> … <<TableEnd:Customers>>).
        doc.MailMerge.ExecuteWithRegions(customers);

        // Save the merged document as a JPEG image.
        // ImageSaveOptions allows us to control rendering; we use the JPEG format.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        // Optional: render only the first page.
        // jpegOptions.PageSet = new PageSet(0);

        doc.Save("MergedResult.jpg", jpegOptions);
    }
}
