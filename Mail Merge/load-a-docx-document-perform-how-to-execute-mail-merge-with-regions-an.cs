using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeWithRegionsToPng
{
    static void Main()
    {
        // Path to the source DOCX file that contains mail‑merge regions.
        const string inputPath = "TemplateWithRegions.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Prepare a DataSet whose table name matches the region name in the document.
        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("CustomerName");
        customers.Columns.Add("Address");
        customers.Rows.Add(new object[] { "Thomas Hardy", "120 Hanover Sq., London" });
        customers.Rows.Add(new object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });

        DataSet data = new DataSet();
        data.Tables.Add(customers);

        // Execute mail merge with regions.
        doc.MailMerge.ExecuteWithRegions(data);

        // Save the merged document as a PNG image.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        const string outputPath = "MergedResult.png";
        doc.Save(outputPath, pngOptions);
    }
}
