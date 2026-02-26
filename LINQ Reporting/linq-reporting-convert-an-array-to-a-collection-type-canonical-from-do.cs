using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOTX template that contains reporting tags.
        Document template = new Document("Template.dotx");

        // Create an array of data objects. The class must have public properties that match the fields used in the template.
        Customer[] customersArray = new Customer[]
        {
            new Customer { FullName = "Thomas Hardy", Address = "120 Hanover Sq., London" },
            new Customer { FullName = "Paolo Accorti", Address = "Via Monte Bianco 34, Torino" }
        };

        // Convert the array to a canonical collection type (List<T>) required by the ReportingEngine.
        List<Customer> customers = customersArray.ToList();

        // Populate the template with the data source. The third argument is the name used in the template to reference the collection.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, customers, "Customers");

        // Save the generated report.
        template.Save("Report.docx");
    }

    // Simple data class whose property names correspond to the merge fields in the DOTX template.
    public class Customer
    {
        public string FullName { get; set; }
        public string Address { get; set; }
    }
}
