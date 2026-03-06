using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    // Simple data entity that matches the merge fields in the template.
    public class Customer
    {
        public string FullName { get; set; }
        public string Address { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOTX template that contains MERGEFIELDs "FullName" and "Address".
            string templatePath = "Template.dotx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Sample collection of customers.
            List<Customer> allCustomers = new List<Customer>
            {
                new Customer { FullName = "Thomas Hardy", Address = "120 Hanover Sq., London" },
                new Customer { FullName = "Paolo Accorti", Address = "Via Monte Bianco 34, Torino" },
                new Customer { FullName = "John Doe", Address = "123 Main St., New York" }
            };

            // LINQ query: select customers whose address contains a comma (just as an example filter).
            var query = from c in allCustomers
                        where c.Address.Contains(",")
                        select c;

            // Convert the LINQ result to a DataTable – this is a format accepted directly by Aspose.Words mail merge.
            DataTable table = new DataTable("Customers");
            table.Columns.Add("FullName", typeof(string));
            table.Columns.Add("Address", typeof(string));

            foreach (Customer cust in query)
            {
                table.Rows.Add(cust.FullName, cust.Address);
            }

            // Perform the mail merge using the DataTable as the data source.
            doc.MailMerge.Execute(table);

            // Save the merged document.
            string outputPath = "Result.docx";
            doc.Save(outputPath);
        }
    }
}
