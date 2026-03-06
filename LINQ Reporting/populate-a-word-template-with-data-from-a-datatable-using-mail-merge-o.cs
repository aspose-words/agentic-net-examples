using System;
using System.Data;
using Aspose.Words;

namespace MailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Path to the DOTX template that contains MERGEFIELDs.
            const string templatePath = @"C:\Templates\MyTemplate.dotx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create a DataTable with column names that match the MERGEFIELD names in the template.
            DataTable table = new DataTable("CustomerData");
            table.Columns.Add("FullName");
            table.Columns.Add("Address");
            table.Columns.Add("City");
            table.Columns.Add("Country");

            // Add sample rows.
            table.Rows.Add("Thomas Hardy", "120 Hanover Sq.", "London", "UK");
            table.Rows.Add("Paolo Accorti", "Via Monte Bianco 34", "Torino", "Italy");

            // Perform the mail merge. All rows will be merged; the document will repeat for each row.
            doc.MailMerge.Execute(table);

            // Save the merged document.
            const string outputPath = @"C:\Output\MergedDocument.docx";
            doc.Save(outputPath);
        }
    }
}
