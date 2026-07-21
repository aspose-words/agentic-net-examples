using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingNestedForeach
{
    // Root wrapper for the data source.
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    // Customer (master) entity.
    public class Customer
    {
        public string Name { get; set; } = string.Empty;
        public List<Invoice> Invoices { get; set; } = new();
    }

    // Invoice (detail) entity.
    public class Invoice
    {
        public int Id { get; set; }
        public decimal Amount { get; set; }
        public DateTime Date { get; set; }
    }

    public class Program
    {
        private const string TemplatePath = "Template.docx";
        private const string OutputPath = "Report.docx";

        public static void Main()
        {
            // 1. Create the template document with nested foreach tags.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Outer foreach for customers.
            builder.Writeln("<<foreach [c in Customers]>>");
            builder.Writeln("Customer: <<[c.Name]>>");
            builder.Writeln("Invoices:");

            // Inner foreach for invoices of the current customer.
            builder.Writeln("<<foreach [i in c.Invoices]>>");
            builder.Writeln("- Id: <<[i.Id]>>, Amount: <<[i.Amount]>>, Date: <<[i.Date]>>");
            builder.Writeln("<</foreach>>"); // End inner foreach.

            builder.Writeln("<</foreach>>"); // End outer foreach.

            // Save the template.
            templateDoc.Save(TemplatePath);

            // 2. Load the template for reporting.
            var doc = new Document(TemplatePath);

            // 3. Prepare sample data.
            var model = new ReportModel
            {
                Customers = new List<Customer>
                {
                    new Customer
                    {
                        Name = "Acme Corp",
                        Invoices = new List<Invoice>
                        {
                            new Invoice { Id = 1001, Amount = 1234.56m, Date = new DateTime(2023, 1, 15) },
                            new Invoice { Id = 1002, Amount = 789.00m, Date = new DateTime(2023, 2, 10) }
                        }
                    },
                    new Customer
                    {
                        Name = "Globex Ltd",
                        Invoices = new List<Invoice>
                        {
                            new Invoice { Id = 2001, Amount = 2500.00m, Date = new DateTime(2023, 3, 5) }
                        }
                    }
                }
            };

            // 4. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options required.
            bool success = engine.BuildReport(doc, model, "model");

            // 5. Save the generated report.
            doc.Save(OutputPath);
        }
    }
}
