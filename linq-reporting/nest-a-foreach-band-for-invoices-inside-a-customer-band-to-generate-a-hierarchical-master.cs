using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create the LINQ Reporting template programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Master band: iterate over customers.
        builder.Writeln("<<foreach [customer in model.Customers]>>");
        builder.Writeln("Customer: <<[customer.Name]>>");
        builder.Writeln("");

        // Detail band: iterate over invoices for each customer.
        builder.Writeln("Invoices:");
        builder.Writeln("<<foreach [invoice in customer.Invoices]>>");
        builder.Writeln("- Invoice #: <<[invoice.Number]>>, Date: <<[invoice.Date]>>, Amount: $<<[invoice.Amount]>>");
        builder.Writeln("<</foreach>>");

        // End of master band.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // Load the template for report generation.
        var template = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Customers = new List<Customer>
            {
                new Customer
                {
                    Name = "Acme Corp",
                    Invoices = new List<Invoice>
                    {
                        new Invoice { Number = "INV-001", Date = "2023-01-15", Amount = 1500.00 },
                        new Invoice { Number = "INV-002", Date = "2023-02-20", Amount = 2750.50 }
                    }
                },
                new Customer
                {
                    Name = "Globex Ltd",
                    Invoices = new List<Invoice>
                    {
                        new Invoice { Number = "INV-101", Date = "2023-03-05", Amount = 3200.75 },
                        new Invoice { Number = "INV-102", Date = "2023-04-12", Amount = 4100.00 },
                        new Invoice { Number = "INV-103", Date = "2023-05-30", Amount = 1250.25 }
                    }
                }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        template.Save(outputPath);
    }
}

// Root wrapper class.
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

// Customer (master) class.
public class Customer
{
    public string Name { get; set; } = "";
    public List<Invoice> Invoices { get; set; } = new();
}

// Invoice (detail) class.
public class Invoice
{
    public string Number { get; set; } = "";
    public string Date { get; set; } = "";
    public double Amount { get; set; }
}
