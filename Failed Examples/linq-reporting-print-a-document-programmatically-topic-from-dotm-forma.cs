// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingPrint
{
    // Simple POCO that represents a data record for the mail‑merge.
    public class Customer
    {
        public string FullName { get; set; }
        public string Company { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public bool IsActive { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the DOTM template that contains merge fields.
            //    The Document(string) constructor is the approved "load" rule.
            Document doc = new Document(@"C:\Templates\ReportTemplate.dotm");

            // 2. Obtain data using LINQ. In a real scenario this could be a database query.
            List<Customer> allCustomers = GetSampleData();

            // Filter only active customers – this demonstrates LINQ usage.
            List<Customer> activeCustomers = allCustomers
                .Where(c => c.IsActive)
                .ToList();

            // 3. Execute mail‑merge. The MailMerge object works with any IEnumerable,
            //    so we can pass the LINQ‑filtered list directly.
            doc.MailMerge.Execute(activeCustomers);

            // 4. (Optional) Update the layout before printing. Required when the document
            //    has been modified after the first render.
            doc.UpdatePageLayout();

            // 5. Print the document programmatically.
            //    The Print() method follows the approved "print" rule.
            //    If a specific printer is required, configure PrinterSettings and use the overload.
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Example: print only the first three pages.
                PrintRange = PrintRange.SomePages,
                FromPage = 1,
                ToPage = 3
            };

            doc.Print(printerSettings);

            // 6. (Optional) Save the generated report for archival purposes.
            //    The Save(string) method is the approved "save" rule.
            doc.Save(@"C:\Reports\GeneratedReport.docx");
        }

        // Helper method that returns a static list of customers.
        // In practice this could be replaced with a LINQ‑to‑SQL/Entity Framework query.
        private static List<Customer> GetSampleData()
        {
            return new List<Customer>
            {
                new Customer { FullName = "Thomas Hardy", Company = "Acme Corp", Address = "120 Hanover Sq.", City = "London", IsActive = true },
                new Customer { FullName = "Paolo Accorti", Company = "Beta Ltd", Address = "Via Monte Bianco 34", City = "Torino", IsActive = false },
                new Customer { FullName = "Jane Doe", Company = "Gamma Inc", Address = "500 Market St.", City = "San Francisco", IsActive = true }
            };
        }
    }
}
