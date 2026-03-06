using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeToPng
{
    static void Main()
    {
        // Load the DOCX template that contains MERGEFIELDs.
        Document template = new Document("Template.docx");

        // Create a simple data source with two records.
        DataTable data = new DataTable();
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Rows.Add("John", "Doe");
        data.Rows.Add("Jane", "Smith");

        // Process each record separately.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the template so each record works on a fresh copy.
            Document doc = (Document)template.Clone(true);

            // Perform a mail merge for the current DataRow.
            doc.MailMerge.Execute(data.Rows[i]);

            // Configure PNG output – render the first page of the merged document.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the first page (zero‑based index).
                PageSet = new PageSet(0)
            };

            // Save the rendered page as a PNG file.
            string outputFile = $"MergedDocument_{i + 1}.png";
            doc.Save(outputFile, pngOptions);
        }
    }
}
