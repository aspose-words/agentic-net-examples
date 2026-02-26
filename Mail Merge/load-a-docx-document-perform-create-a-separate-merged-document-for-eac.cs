using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeToPng
{
    static void Main()
    {
        // Path to the mail‑merge template document (must contain MERGEFIELDs).
        const string templatePath = "Template.docx";

        // Load the template document.
        Document template = new Document(templatePath);

        // Prepare a data source – a DataTable with one row per merged document.
        DataTable data = new DataTable("Customers");
        data.Columns.Add("Name");
        data.Columns.Add("Address");
        data.Rows.Add("John Doe", "123 Main St, Anytown");
        data.Rows.Add("Jane Smith", "456 Oak Ave, Othertown");
        data.Rows.Add("Bob Johnson", "789 Pine Rd, Sometown");

        // Iterate over each data row, create a merged document and save it as PNG.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the template so that each iteration works with a fresh copy.
            Document mergedDoc = (Document)template.Clone(true);

            // Execute mail merge for the current row only.
            mergedDoc.MailMerge.Execute(data.Rows[i]);

            // Define the output PNG file name.
            string outputPath = $"MergedDocument_{i + 1}.png";

            // Configure image save options – render the first page as PNG.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the first page (zero‑based index).
                PageSet = new PageSet(0)
            };

            // Save the merged document page as a PNG image.
            mergedDoc.Save(outputPath, pngOptions);
        }
    }
}
