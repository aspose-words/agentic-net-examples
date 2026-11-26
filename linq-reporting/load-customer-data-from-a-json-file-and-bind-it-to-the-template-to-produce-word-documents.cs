using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace LinqReportingExample
{
    // Data models
    public class Customer
    {
        public string Name { get; set; } = "";
        public string Email { get; set; } = "";
        public string Phone { get; set; } = "";
        public Address Address { get; set; } = new();
    }

    public class Address
    {
        public string Street { get; set; } = "";
        public string City { get; set; } = "";
        public string Country { get; set; } = "";
    }

    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required by Aspose.Words for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample JSON data
            string jsonPath = "customers.json";
            var sampleData = new ReportModel
            {
                Customers = new List<Customer>
                {
                    new Customer
                    {
                        Name = "Alice Johnson",
                        Email = "alice@example.com",
                        Phone = "+1-555-0100",
                        Address = new Address
                        {
                            Street = "123 Maple St",
                            City = "Springfield",
                            Country = "USA"
                        }
                    },
                    new Customer
                    {
                        Name = "Bob Smith",
                        Email = "bob@example.com",
                        Phone = "+1-555-0200",
                        Address = new Address
                        {
                            Street = "456 Oak Ave",
                            City = "Rivertown",
                            Country = "USA"
                        }
                    }
                }
            };
            File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

            // Load JSON data into model
            string jsonContent = File.ReadAllText(jsonPath);
            var model = JsonConvert.DeserializeObject<ReportModel>(jsonContent) ?? new ReportModel();

            // Create template document with LINQ Reporting tags
            string templatePath = "template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Customer Report");
            builder.Writeln();

            // Begin foreach block
            builder.Writeln("<<foreach [c in Customers]>>");

            // Table definition inside foreach
            Table table = builder.StartTable();

            // Header row
            builder.InsertCell(); builder.Writeln("Name");
            builder.InsertCell(); builder.Writeln("Email");
            builder.InsertCell(); builder.Writeln("Phone");
            builder.InsertCell(); builder.Writeln("City");
            builder.EndRow();

            // Data row (will be repeated for each customer)
            builder.InsertCell(); builder.Writeln("<<[c.Name]>>");
            builder.InsertCell(); builder.Writeln("<<[c.Email]>>");
            builder.InsertCell(); builder.Writeln("<<[c.Phone]>>");
            builder.InsertCell(); builder.Writeln("<<[c.Address.City]>>");
            builder.EndRow();

            // End of table
            builder.EndTable();

            // End foreach block
            builder.Writeln("<</foreach>>");

            // Save the template
            templateDoc.Save(templatePath);

            // Load the template for reporting
            var reportDoc = new Document(templatePath);

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report
            string outputPath = "CustomerReport.docx";
            reportDoc.Save(outputPath);
        }
    }
}
