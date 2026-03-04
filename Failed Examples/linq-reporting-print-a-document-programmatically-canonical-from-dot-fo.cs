// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using Aspose.Words;

namespace AsposeWordsLinqReportingPrint
{
    // Simple POCO class representing a customer.
    public class Customer
    {
        public string Name { get; set; }
        public string Address { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the DOT template that contains merge fields "Name" and "Address".
            //    The Document constructor with a file name loads the file.
            Document template = new Document(@"C:\Templates\ReportTemplate.dot");

            // 2. Prepare a data source. In a real scenario this could come from a database.
            List<Customer> customers = new List<Customer>
            {
                new Customer { Name = "James Bond", Address = "MI5 Headquarters, London" },
                new Customer { Name = "Ethan Hunt", Address = "Impossible Missions Force, USA" },
                new Customer { Name = "Natasha Romanoff", Address = "Avengers Tower, New York" }
            };

            // 3. Use LINQ to select the first customer whose name starts with 'J'.
            Customer selected = customers
                .Where(c => c.Name.StartsWith("J"))
                .FirstOrDefault();

            if (selected == null)
                throw new InvalidOperationException("No matching customer found.");

            // 4. Execute mail merge. The field names must match the property names of the data object.
            //    MailMerge.Execute takes an object; Aspose.Words will reflect over its public properties.
            template.MailMerge.Execute(selected);

            // 5. Optionally update fields (e.g., DATE fields) before printing.
            template.UpdateFields();

            // 6. Print the document using the default printer.
            //    The Print() method prints the whole document without UI.
            template.Print();

            // 7. If you need to print to a specific printer or a page range, configure PrinterSettings.
            //    Below is an example of printing pages 1‑2 to a named printer.
            /*
            PrinterSettings settings = new PrinterSettings
            {
                PrinterName = "Microsoft Print to PDF",
                PrintRange = PrintRange.SomePages,
                FromPage = 1,
                ToPage = 2
            };
            template.Print(settings);
            */

            // 8. (Optional) Save the merged document for archival purposes.
            //    Save format is inferred from the extension; here we save as DOCX.
            template.Save(@"C:\Output\ReportResult.docx");
        }
    }
}
