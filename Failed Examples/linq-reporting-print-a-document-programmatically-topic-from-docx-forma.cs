// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains MERGEFIELDs (e.g., «Name», «Address»)
            const string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document (lifecycle rule: load)
            Document doc = new Document(templatePath);

            // Sample data source – a list of customers
            List<Customer> customers = new List<Customer>
            {
                new Customer { Name = "Alice Johnson",   Address = "123 Maple St, Springfield" },
                new Customer { Name = "Bob Smith",       Address = "456 Oak Ave, Shelbyville" },
                new Customer { Name = "Charlie Daniels", Address = "789 Pine Rd, Capital City" }
            };

            // Use LINQ to select only customers from a specific city (example filter)
            var filteredCustomers = customers
                .Where(c => c.Address.Contains("Springfield") || c.Address.Contains("Shelbyville"))
                .ToArray();

            // Perform mail merge using the filtered data (feature: mail merge)
            // The field names must match the MERGEFIELD names in the template.
            doc.MailMerge.Execute(
                new string[] { "Name", "Address" },
                filteredCustomers.Select(c => new object[] { c.Name, c.Address }).SelectMany(a => a).ToArray()
            );

            // Optional: update fields (e.g., DATE fields) before printing
            doc.UpdateFields();

            // Print the document to the default printer (feature: print)
            // You can also specify a printer name or PrinterSettings if needed.
            doc.Print();

            // If you need to print to a specific printer, uncomment the following lines:
            // PrinterSettings settings = new PrinterSettings();
            // settings.PrinterName = "Your Printer Name";
            // doc.Print(settings);
        }
    }

    // Simple POCO class representing a customer record
    public class Customer
    {
        public string Name { get; set; }
        public string Address { get; set; }
    }
}
