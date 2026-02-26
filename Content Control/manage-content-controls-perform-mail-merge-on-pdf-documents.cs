using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsPdfMailMerge
{
    class Program
    {
        static void Main()
        {
            // Path to the source PDF that contains MERGEFIELD tags.
            const string inputPdfPath = @"C:\Docs\Template.pdf";

            // Load the PDF as a Word document. Aspose.Words can read PDF files when the LoadFormat is set to Pdf.
            var loadOptions = new LoadOptions { LoadFormat = LoadFormat.Pdf };
            Document doc = new Document(inputPdfPath, loadOptions);

            // Prepare a DataTable that matches the merge fields present in the PDF.
            // Example merge fields: FirstName, LastName, Address
            DataTable data = new DataTable("MailMergeData");
            data.Columns.Add("FirstName");
            data.Columns.Add("LastName");
            data.Columns.Add("Address");

            // Add rows of data – each row will produce a separate merged document page.
            data.Rows.Add("John", "Doe", "123 Main St, Anytown");
            data.Rows.Add("Jane", "Smith", "456 Oak Ave, Othertown");

            // Perform the mail merge. This replaces the MERGEFIELD tags with the values from the DataTable.
            doc.MailMerge.Execute(data);

            // Save the merged result back to PDF. PdfSaveOptions can be used to control PDF output if needed.
            var pdfSaveOptions = new PdfSaveOptions
            {
                // Example: keep the original document title in the PDF viewer window.
                DisplayDocTitle = true
            };

            const string outputPdfPath = @"C:\Docs\MergedResult.pdf";
            doc.Save(outputPdfPath, pdfSaveOptions);
        }
    }
}
