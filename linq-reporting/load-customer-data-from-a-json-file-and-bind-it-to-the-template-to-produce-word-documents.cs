using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingExample
{
    // Data model classes
    public class Customer
    {
        public string Name { get; set; } = "";
        public string Email { get; set; } = "";
        public List<Order> Orders { get; set; } = new();
    }

    public class Order
    {
        public string Product { get; set; } = "";
        public int Quantity { get; set; }
    }

    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words if needed
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create sample JSON data file
            const string jsonFileName = "customers.json";
            var sampleData = new ReportModel
            {
                Customers = new List<Customer>
                {
                    new Customer
                    {
                        Name = "Alice Johnson",
                        Email = "alice@example.com",
                        Orders = new List<Order>
                        {
                            new Order { Product = "Laptop", Quantity = 1 },
                            new Order { Product = "Mouse", Quantity = 2 }
                        }
                    },
                    new Customer
                    {
                        Name = "Bob Smith",
                        Email = "bob@example.com",
                        Orders = new List<Order>
                        {
                            new Order { Product = "Keyboard", Quantity = 1 }
                        }
                    }
                }
            };
            File.WriteAllText(jsonFileName, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

            // Create the LINQ Reporting template programmatically
            const string templateFileName = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Customer Report");
            builder.Writeln("<<foreach [c in Customers]>>");
            builder.Writeln("Name: <<[c.Name]>>");
            builder.Writeln("Email: <<[c.Email]>>");
            builder.Writeln("Orders:");
            builder.Writeln("<<foreach [o in c.Orders]>>");
            builder.Writeln("- <<[o.Product]>> (Qty: <<[o.Quantity]>>)");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templateFileName);

            // Load JSON data into the model
            var jsonContent = File.ReadAllText(jsonFileName);
            var model = JsonConvert.DeserializeObject<ReportModel>(jsonContent) ?? new ReportModel();

            // Load the template document
            var doc = new Document(templateFileName);

            // Build the report using LINQ Reporting Engine
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            const string outputFileName = "Report.docx";
            doc.Save(outputFileName);
        }
    }
}
