using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeToPng
{
    static void Main()
    {
        // Path to the mail‑merge template (must contain MERGEFIELDs).
        const string templatePath = @"C:\Templates\ReportTemplate.docx";

        // Load the template document.
        Document template = new Document(templatePath);

        // Prepare a data source – a DataTable with two rows as an example.
        DataTable data = new DataTable("ReportData");
        data.Columns.Add("Title");
        data.Columns.Add("Value");
        data.Rows.Add("First", 123);
        data.Rows.Add("Second", 456);

        // Iterate over each row, create a merged document and save it as PNG.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the template to keep the original unchanged.
            Document merged = (Document)template.Clone(true);

            // Execute mail merge for the current row only.
            merged.MailMerge.Execute(data.Rows[i]);

            // Define the output file name.
            string outputPath = $@"C:\Output\MergedDocument_{i + 1}.png";

            // Save the merged document as a PNG image (first page rendered).
            merged.Save(outputPath, SaveFormat.Png);
        }
    }
}
