// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load the DOTM template (the document that contains merge fields).
        Document template = new Document("Template.dotm");

        // --------------------------------------------------------------------
        // Prepare a LINQ data source.
        // --------------------------------------------------------------------
        var people = new[]
        {
            new Person
            {
                FullName = "John Doe",
                Company  = "Acme Corp",
                Address  = "123 Main St",
                City     = "Metropolis"
            },
            new Person
            {
                FullName = "Jane Smith",
                Company  = "Globex Inc",
                Address  = "456 Oak Ave",
                City     = "Gotham"
            }
        };

        // Convert the LINQ result to a DataTable – the format required by MailMerge.
        DataTable dataTable = new DataTable();
        dataTable.Columns.Add("FullName");
        dataTable.Columns.Add("Company");
        dataTable.Columns.Add("Address");
        dataTable.Columns.Add("City");

        foreach (var p in people)
        {
            dataTable.Rows.Add(p.FullName, p.Company, p.Address, p.City);
        }

        // --------------------------------------------------------------------
        // Perform mail merge for each record, print and optionally save.
        // --------------------------------------------------------------------
        foreach (DataRow row in dataTable.Rows)
        {
            // Clone the template so each iteration works with a fresh copy.
            Document mergedDoc = (Document)template.Clone(true);

            // Execute mail merge using the column names as field names.
            mergedDoc.MailMerge.Execute(
                dataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray(),
                row.ItemArray);

            // Re‑calculate page layout – required before printing.
            mergedDoc.UpdatePageLayout();

            // Print the merged document to the default printer.
            mergedDoc.Print();

            // Save the merged document (optional).
            string outFile = $"Output_{row["FullName"]}.docx";
            mergedDoc.Save(outFile);
        }
    }

    // Simple POCO used for the LINQ query.
    class Person
    {
        public string FullName { get; set; }
        public string Company  { get; set; }
        public string Address  { get; set; }
        public string City     { get; set; }
    }
}
