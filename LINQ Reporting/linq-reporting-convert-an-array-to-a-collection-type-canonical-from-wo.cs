using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // ---------- Create a new blank document ----------
        Document doc = new Document();

        // ---------- Load a WORDML (DOCX) template ----------
        // The template should contain LINQ Reporting tags, e.g. <<foreach [Customers]>><<[FullName]>><</foreach>>
        doc = new Document("Template.docx");

        // ---------- Sample data as an array ----------
        Customer[] customerArray = new[]
        {
            new Customer { FullName = "Thomas Hardy", Address = "120 Hanover Sq., London" },
            new Customer { FullName = "Paolo Accorti", Address = "Via Monte Bianco 34, Torino" }
        };

        // ---------- Convert the array to a canonical collection type ----------
        // ReportingEngine works with IList, IEnumerable, etc. Converting to List<T> is the canonical approach.
        List<Customer> customerList = customerArray.ToList();

        // ---------- (Optional) Demonstrate converting a collection to an array ----------
        // For example, get all sections of the document and copy them to an array.
        Section[] sectionsArray = doc.Sections.ToArray();

        // ---------- Build the report using the collection ----------
        ReportingEngine engine = new ReportingEngine();
        // The third argument is the name used inside the template to reference this data source.
        engine.BuildReport(doc, customerList, "Customers");

        // ---------- Save the populated document ----------
        doc.Save("Report.docx");
    }

    // Simple POCO class used as a data source for the report.
    public class Customer
    {
        public string FullName { get; set; }
        public string Address { get; set; }
    }
}
