using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used as the data source for the report.
    public class Order
    {
        public int Id { get; set; }
        public double Amount { get; set; }
        public DateTime Date { get; set; }
    }

    public class Customer
    {
        public string Name { get; set; }
        public List<Order> Orders { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document in memory.
            // -----------------------------------------------------------------
            Document template = new Document();                     // create document
            DocumentBuilder builder = new DocumentBuilder(template);

            // Write a heading.
            builder.Writeln("Customer Report");
            builder.Writeln();

            // Insert a field that references a property of the root object.
            // The template syntax <<[customer.Name]>> will be replaced by the customer's name.
            builder.Writeln("Customer: <<[customer.Name]>>");
            builder.Writeln();

            // Insert a table header for the orders.
            builder.Writeln("Orders:");
            builder.Writeln("Id\tAmount\tDate");
            // Use a foreach loop to iterate over the collection.
            // Inside the loop we can also call static members, e.g. Math.Round.
            builder.Writeln("<<foreach [in customer.Orders]>>");
            builder.Writeln("<<[Id]>>\t<<[Amount]:dollarText>>\t<<[Date].ToString(\"yyyy-MM-dd\")>>");
            builder.Writeln("<</foreach>>");

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            var customer = new Customer
            {
                Name = "John Doe",
                Orders = new List<Order>
                {
                    new Order { Id = 1, Amount = 1234.56, Date = new DateTime(2023, 1, 15) },
                    new Order { Id = 2, Amount = 789.00, Date = new DateTime(2023, 2, 20) },
                    new Order { Id = 3, Amount = 456.78, Date = new DateTime(2023, 3, 5) }
                }
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting Engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Add any types whose static members we want to use inside the template.
            // For this example we add System.Math so that we could call Math.Round if needed.
            engine.KnownTypes.Add(typeof(Math));

            // BuildReport overload that allows referencing the data source object itself.
            // The third argument ("customer") is the name used in the template.
            engine.BuildReport(template, customer, "customer");

            // -----------------------------------------------------------------
            // 4. Save the populated document.
            // -----------------------------------------------------------------
            template.Save("CustomerReport.docx"); // save document
        }
    }
}
