// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load the DOCM template that contains merge fields.
        Document doc = new Document("ReportTemplate.docm");

        // -----------------------------------------------------------------
        // Example data source – a list of customers.
        // -----------------------------------------------------------------
        var customers = new List<Customer>
        {
            new Customer { Name = "Alice",   City = "London",  IsActive = true  },
            new Customer { Name = "Bob",     City = "Paris",   IsActive = false },
            new Customer { Name = "Charlie", City = "Berlin",  IsActive = true  }
        };

        // -----------------------------------------------------------------
        // Use LINQ to select only the active customers.
        // -----------------------------------------------------------------
        var activeCustomers = customers
            .Where(c => c.IsActive)
            .Select(c => new { c.Name, c.City })
            .ToArray();

        // -----------------------------------------------------------------
        // Prepare field names and values for the mail‑merge operation.
        // -----------------------------------------------------------------
        string[] fieldNames = { "Name", "City" };
        object[] fieldValues = activeCustomers
            .SelectMany(c => new object[] { c.Name, c.City })
            .ToArray();

        // Execute the mail merge – the template will be populated with the data.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Rebuild the page layout so that pagination is correct before printing.
        doc.UpdatePageLayout();

        // -----------------------------------------------------------------
        // Print the document to the default printer.
        // -----------------------------------------------------------------
        doc.Print();

        // -----------------------------------------------------------------
        // Optional: print with custom printer settings.
        // -----------------------------------------------------------------
        //PrinterSettings printerSettings = new PrinterSettings();
        //printerSettings.PrinterName = "Your Printer Name";
        //printerSettings.PrintRange = PrintRange.AllPages;
        //doc.Print(printerSettings);
    }
}

// Simple POCO representing a customer – used only for the LINQ example.
public class Customer
{
    public string Name { get; set; }
    public string City { get; set; }
    public bool   IsActive { get; set; }
}
