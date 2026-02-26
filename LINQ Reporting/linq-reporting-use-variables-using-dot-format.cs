using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new empty document and a builder to add template tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use DOT notation in LINQ Reporting tags to access nested members.
        builder.Writeln("Customer: <<[order.Customer.Name]>>");
        builder.Writeln("City: <<[order.Customer.Address.City]>>");
        builder.Writeln("Order total: <<[order.Total]>>");

        // Prepare a data source with nested objects.
        var data = new
        {
            order = new Order
            {
                Total = 123.45,
                Customer = new Customer
                {
                    Name = "John Doe",
                    Address = new Address
                    {
                        City = "London",
                        Street = "123 Main St."
                    }
                }
            }
        };

        // Build the report. The third argument gives a name to reference the data source itself.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "order");

        // Save the populated document.
        doc.Save("Report.docx");
    }

    // POCO classes used as the data source.
    public class Order
    {
        public double Total { get; set; }
        public Customer Customer { get; set; }
    }

    public class Customer
    {
        public string Name { get; set; }
        public Address Address { get; set; }
    }

    public class Address
    {
        public string City { get; set; }
        public string Street { get; set; }
    }
}
