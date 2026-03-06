using System;
using System.Data;
using Aspose.Words;

namespace MailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Path to the mail‑merge template (must contain MERGEFIELDs with the same names as the data source).
            const string templatePath = "Template.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Example 1: Simple array‑based data source (single record).
            string[] fieldNames = { "FullName", "Address", "City" };
            object[] fieldValues = { "James Bond", "MI5 Headquarters", "London" };
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Example 2: DataTable data source (multiple records).
            // Uncomment the following block if you need to merge several rows.
            /*
            DataTable table = new DataTable("Customers");
            table.Columns.Add("FullName");
            table.Columns.Add("Address");
            table.Columns.Add("City");
            table.Rows.Add("John Doe", "123 Main St.", "New York");
            table.Rows.Add("Jane Smith", "456 Oak Ave.", "Los Angeles");
            doc.MailMerge.Execute(table);
            */

            // Save the merged document.
            const string outputPath = "Merged.docx";
            doc.Save(outputPath);
        }
    }
}
