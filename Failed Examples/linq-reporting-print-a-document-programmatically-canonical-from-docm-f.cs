// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCM template from the file system.
        Document doc = new Document("ReportTemplate.docm");

        // Retrieve the data that will be merged into the document.
        List<Customer> customers = GetCustomers();

        // Use LINQ to filter only customers that have at least one order.
        var filtered = customers.Where(c => c.OrderCount > 0).ToList();

        // Prepare the field names defined in the DOCM template.
        string[] fieldNames = new[] { "Name", "OrderCount", "Total" };

        // Concatenate the names, sum the order counts and totals for the filtered set.
        object[] fieldValues = new object[]
        {
            string.Join(", ", filtered.Select(c => c.Name)),
            filtered.Sum(c => c.OrderCount),
            filtered.Sum(c => c.Total)
        };

        // Perform the mail merge to populate the template.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Update any fields (e.g., DATE, PAGE) and rebuild the page layout.
        doc.UpdateFields();
        doc.UpdatePageLayout();

        // Set up printer settings – print the whole document.
        PrinterSettings printerSettings = new PrinterSettings
        {
            PrintRange = PrintRange.AllPages
        };

        // Print the document using the configured printer settings.
        doc.Print(printerSettings);
    }

    // Mock method that returns a collection of customers.
    static List<Customer> GetCustomers()
    {
        return new List<Customer>
        {
            new Customer { Name = "Alice",   OrderCount = 3, Total = 150.0 },
            new Customer { Name = "Bob",     OrderCount = 0, Total =   0.0 },
            new Customer { Name = "Charlie", OrderCount = 5, Total = 300.0 }
        };
    }

    // Simple POCO representing a customer record.
    class Customer
    {
        public string Name { get; set; }
        public int    OrderCount { get; set; }
        public double Total { get; set; }
    }
}
