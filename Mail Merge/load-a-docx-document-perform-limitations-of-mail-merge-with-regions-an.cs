using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Create a DataTable that will serve as the mail‑merge data source.
        // The table name must match the region name defined in the document
        // (e.g. MERGEFIELD TableStart:RegionName / TableEnd:RegionName).
        DataTable table = new DataTable("Region");
        table.Columns.Add("Name");
        table.Columns.Add("Address");
        table.Rows.Add("John Doe", "123 Main St");
        table.Rows.Add("Jane Smith", "456 Oak Ave");

        // Perform mail merge using regions.
        doc.MailMerge.ExecuteWithRegions(table);

        // Configure image save options to render the document as JPEG.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        // Render the first page (zero‑based index). Adjust PageSet if you need other pages.
        jpegOptions.PageSet = new PageSet(0);

        // Save the merged document as a JPEG image.
        doc.Save("Output.jpg", jpegOptions);
    }
}
