using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // 1. Create a template document that contains LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // The tags use the data source name "customer" to access its members.
        builder.Writeln("Customer: <<[customer.FullName]>>");
        builder.Writeln("Address: <<[customer.Address]>>");

        // 2. Prepare the data source object.
        var customer = new Customer
        {
            FullName = "John Doe",
            Address = "123 Main St, Anytown"
        };

        // 3. Build the report, allowing the template to reference the data source object itself.
        ReportingEngine engine = new ReportingEngine();
        // Use the overload that accepts a data source name.
        engine.BuildReport(template, customer, "customer");

        // 4. Save the populated document as PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions(); // default options
        template.Save("Report.pdf", pdfOptions);
    }

    // Simple POCO used as the data source.
    public class Customer
    {
        public string FullName { get; set; }
        public string Address { get; set; }
    }
}
