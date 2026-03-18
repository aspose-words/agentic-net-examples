using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMailMergeToPdf
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX template that contains MERGEFIELDs.
            string templatePath = @"C:\Docs\Template.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a simple DataTable with data for the mail merge.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("FullName");
            table.Columns.Add("Company");
            table.Columns.Add("Address");
            table.Columns.Add("City");

            table.Rows.Add("James Bond", "MI5 Headquarters", "Milbank", "London");
            table.Rows.Add("Ethan Hunt", "IMF", "123 Secret St.", "Washington");

            // Execute the mail merge using the DataTable.
            doc.MailMerge.Execute(table);

            // Save the merged document as PDF.
            // Use the overload that accepts a file name and a SaveFormat enum.
            string pdfOutputPath = @"C:\Docs\MergedResult.pdf";
            doc.Save(pdfOutputPath, SaveFormat.Pdf);

            Console.WriteLine("Mail merge completed and PDF saved to: " + pdfOutputPath);
        }
    }
}
