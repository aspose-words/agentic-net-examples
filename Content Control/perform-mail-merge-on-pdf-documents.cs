using System;
using System.Data;
using Aspose.Words;

namespace MailMergePdfExample
{
    class Program
    {
        static void Main()
        {
            // Load the PDF template that contains MERGEFIELD fields.
            // Aspose.Words can open PDF files and treat them as a Document object.
            Document pdfTemplate = new Document("Template.pdf");

            // Prepare a data source. The column names must match the merge field names in the PDF.
            DataTable data = new DataTable("MailData");
            data.Columns.Add("FirstName");
            data.Columns.Add("LastName");
            data.Columns.Add("Address");

            // Add sample rows.
            data.Rows.Add("James", "Bond", "MI5 Headquarters, London");
            data.Rows.Add("Ethan", "Hunt", "Impossible Missions Force, USA");

            // Execute the mail merge. All rows will be merged; the document will be expanded as needed.
            pdfTemplate.MailMerge.Execute(data);

            // Save the merged result back to PDF.
            pdfTemplate.Save("MergedResult.pdf");
        }
    }
}
