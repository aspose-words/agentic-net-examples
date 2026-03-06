using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MailMergeToPdfExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX template that contains MERGEFIELDs.
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the DOCX document using the Document constructor (lifecycle rule).
            Document doc = new Document(templatePath);

            // Prepare a simple data source – a DataTable with column names matching the merge fields.
            DataTable data = new DataTable("ReportData");
            data.Columns.Add("Title");
            data.Columns.Add("Author");
            data.Columns.Add("Date");
            data.Columns.Add("Content");

            // Add a single row of data for the report.
            data.Rows.Add(
                "Quarterly Financial Summary",
                "Jane Doe",
                DateTime.Now.ToString("MMMM dd, yyyy"),
                "The company achieved a 12% increase in revenue compared to the previous quarter."
            );

            // Execute the mail merge for the single record (using the MailMerge.Execute(DataTable) method).
            doc.MailMerge.Execute(data);

            // Convert the merged document to PDF.
            // Save method overload with file name and SaveFormat (lifecycle rule) is used.
            string outputPdfPath = @"C:\Output\Report.pdf";
            doc.Save(outputPdfPath, SaveFormat.Pdf);

            // Optional: inform the user that the process completed.
            Console.WriteLine("Mail merge completed and PDF saved to: " + outputPdfPath);
        }
    }
}
