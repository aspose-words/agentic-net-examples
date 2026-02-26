using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeToPng
{
    static void Main()
    {
        // Path to the DOCX template that contains MERGEFIELDs.
        string templatePath = "Template.docx";

        // Load the template document.
        Document template = new Document(templatePath);

        // Prepare sample data for the mail merge.
        DataTable data = new DataTable();
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Rows.Add("John", "Doe");
        data.Rows.Add("Jane", "Smith");
        data.Rows.Add("Bob", "Johnson");

        // Iterate through each record, clone the template, execute the merge and save as PNG.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the template so the original remains unchanged for the next iteration.
            Document mergedDoc = (Document)template.Clone(true);

            // Execute a simple mail merge for the current DataRow.
            mergedDoc.MailMerge.Execute(data.Rows[i]);

            // Configure image save options for PNG format.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the first page of the merged document.
                PageSet = new PageSet(0),

                // Optional: set resolution (dpi) for higher quality.
                Resolution = 300
            };

            // Save the merged document page as a PNG image.
            string outputPath = $"MergedDocument_{i + 1}.png";
            mergedDoc.Save(outputPath, pngOptions);
        }
    }
}
