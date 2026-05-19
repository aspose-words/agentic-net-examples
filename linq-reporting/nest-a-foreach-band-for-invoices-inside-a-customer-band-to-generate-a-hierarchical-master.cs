using System;
using System.Collections.Generic;
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
                        new Invoice { Id = 2001, Date = new DateTime(2023, 3, 5), Amount = 2500.00m }
                    }
                }
            }
        };

        // Create a template document with nested foreach bands.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Customer Report");
        builder.Writeln();

        // Customer band.
        builder.Writeln("<<foreach [customer in Customers]>>");
        builder.Writeln("Customer: <<[customer.Name]>>");
        builder.Writeln("Invoices:");

        // Invoice band nested inside the customer band.
        builder.Writeln("<<foreach [invoice in customer.Invoices]>>");
        builder.Writeln("- Id: <<[invoice.Id]>>, Date: <<[invoice.Date]>>, Amount: <<[invoice.Amount]>>");
        builder.Writeln("<</foreach>>"); // end invoice foreach

        builder.Writeln("<</foreach>>"); // end customer foreach

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        template.Save("Report.docx");
    }
}

// Root wrapper for the report data.
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
