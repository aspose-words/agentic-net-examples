using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingNestedBands
{
    // Data model for the master‑detail report.
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Customer
    {
        public string Name { get; set; } = string.Empty;
        public List<Invoice> Invoices { get; set; } = new();
    }

    public class Invoice
    {
        public int Id { get; set; }
        public decimal Amount { get; set; }
        public DateTime Date { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample data.
            var customers = new List<Customer>
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
            };

            // 2. Create the LINQ Reporting template programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Master‑Detail Report");
            builder.Writeln();

            // Outer foreach band for customers.
            builder.Writeln("<<foreach [customer in Customers]>>");
            builder.Writeln("Customer: <<[customer.Name]>>");
            builder.Writeln();

            // Inner foreach band for invoices belonging to the current customer.
            builder.Writeln("  <<foreach [invoice in customer.Invoices]>>");
            builder.Writeln("  Invoice ID: <<[invoice.Id]>>");
            builder.Writeln("  Amount: $<<[invoice.Amount]>>");
            builder.Writeln("  Date: <<[invoice.Date]>>");
            builder.Writeln("  <</foreach>>");
            builder.Writeln();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            doc.Save(templatePath);

            // 3. Load the template and build the report.
            var templateDoc = new Document(templatePath);
            var model = new ReportModel { Customers = customers };
            var engine = new ReportingEngine();

            // Build the report using the model as the root data source.
            engine.BuildReport(templateDoc, model);

            // 4. Save the generated report.
            var outputPath = "Report.docx";
            templateDoc.Save(outputPath);
        }
    }
}
