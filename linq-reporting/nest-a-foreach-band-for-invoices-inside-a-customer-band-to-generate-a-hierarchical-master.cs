using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
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
                        new Invoice { Id = 1001, Date = new DateTime(2023, 1, 15), Amount = 1234.56m },
                        new Invoice { Id = 1002, Date = new DateTime(2023, 2, 20), Amount = 789.00m }
                    }
                },
                new Customer
                {
                    Name = "Globex Ltd",
                    Invoices = new List<Invoice>
                    {
                        new Invoice { Id = 2001, Date = new DateTime(2023, 3, 5), Amount = 456.78m }
                    }
                }
            }
        };

        // Create the template document programmatically.
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln("Customer: <<[c.Name]>>");
        builder.Writeln("");
        builder.Writeln("Invoices:");
        builder.Writeln("<<foreach [i in c.Invoices]>>");
        builder.Writeln("  • Invoice #: <<[i.Id]>>   Date: <<[i.Date]>>   Amount: $<<[i.Amount]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template (optional – we can reuse the same Document instance).
        var template = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        template.Save(outputPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Wrapper model for the report.
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

// Customer data model.
public class Customer
{
    public string Name { get; set; } = string.Empty;
    public List<Invoice> Invoices { get; set; } = new();
}

// Invoice data model.
public class Invoice
{
    public int Id { get; set; }
    public DateTime Date { get; set; }
    public decimal Amount { get; set; }
}
