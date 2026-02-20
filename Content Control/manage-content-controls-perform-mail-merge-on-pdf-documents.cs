using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.MailMerging; // Added for MailMergeCleanupOptions

class Program
{
    static void Main()
    {
        // Path to the PDF template that contains MERGEFIELD fields.
        string templatePath = "Template.pdf";

        // Path where the merged PDF will be saved.
        string outputPath = "MergedResult.pdf";

        // Load the PDF as a Word document. Aspose.Words can open PDF files for reading.
        LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Pdf };
        Document doc = new Document(templatePath, loadOptions);

        // Create a data source for the mail merge.
        DataTable data = new DataTable("Customer");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Address");
        data.Rows.Add("John", "Doe", "120 Hanover Sq., London");
        data.Rows.Add("Jane", "Smith", "123 Main St., New York");

        // Optional: clean up empty paragraphs after the merge.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;

        // Perform the mail merge using the DataTable.
        doc.MailMerge.Execute(data);

        // Save the result back to PDF.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
