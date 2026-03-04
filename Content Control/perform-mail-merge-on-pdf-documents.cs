using System;
using System.Data;
using Aspose.Words;

namespace MailMergePdfExample
{
    class Program
    {
        static void Main()
        {
            // Path to the PDF template that contains MERGEFIELD fields.
            const string templatePath = @"C:\Templates\InvoiceTemplate.pdf";

            // Path where the merged PDF will be saved.
            const string outputPath = @"C:\Output\Invoice_Merged.pdf";

            // Load the PDF document. Aspose.Words can treat PDF as a Word document.
            Document doc = new Document(templatePath);

            // Prepare a DataTable that matches the merge fields in the PDF.
            DataTable data = new DataTable("Invoice");
            data.Columns.Add("CustomerName");
            data.Columns.Add("InvoiceDate");
            data.Columns.Add("TotalAmount");

            // Add a single row of data (mail merge works with one or many rows).
            data.Rows.Add("Acme Corp.", DateTime.Today.ToShortDateString(), "$1,250.00");

            // Perform the mail merge. This will replace MERGEFIELDs with the values from the DataTable.
            doc.MailMerge.Execute(data);

            // Save the result as a PDF document.
            doc.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
