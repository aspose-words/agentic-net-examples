using System;
using System.Data;
using Aspose.Words;

namespace MailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX template that contains MERGEFIELDs.
            // Example template should have fields: CustomerName, Address
            string templatePath = @"C:\Docs\Template.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a DataTable that matches the merge fields in the template.
            DataTable table = new DataTable("Customers");
            table.Columns.Add("CustomerName");
            table.Columns.Add("Address");

            // Add rows of data.
            table.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");
            table.Rows.Add("Paolo Accorti", "Via Monte Bianco 34, Torino");

            // Execute the mail merge. All rows will be merged, producing a document
            // that repeats the whole template for each record.
            doc.MailMerge.Execute(table);

            // Save the merged document.
            string outputPath = @"C:\Docs\MergedResult.docx";
            doc.Save(outputPath);
        }
    }
}
