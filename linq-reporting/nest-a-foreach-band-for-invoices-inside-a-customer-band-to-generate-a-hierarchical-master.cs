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
                        new Invoice { Id = 1001, Amount = 1234.56m },
                        new Invoice { Id = 1002, Amount = 789.00m }
                    }
                },
                new Customer
                {
                    Name = "Globex Ltd",
                    Invoices = new List<Invoice>
                    {
                        new Invoice { Id = 2001, Amount = 456.78m }
                    }
                }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customers Report");
        builder.Writeln("<<foreach [customer in Customers]>>");
        builder.Writeln("Customer: <<[customer.Name]>>");
        builder.Writeln("Invoices:");
        builder.Writeln("<<foreach [invoice in customer.Invoices]>>");
        builder.Writeln("- Invoice ID: <<[invoice.Id]>>  Amount: <<[invoice.Amount]>>");
        builder.Writeln("<</foreach>>"); // end inner foreach (invoices)
        builder.Writeln("<</foreach>>"); // end outer foreach (customers)

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Root wrapper class for the report data.
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

// Customer (master) class.
public class Customer
{
    public string Name { get; set; } = string.Empty;
    public List<Invoice> Invoices { get; set; } = new();
}

// Invoice (detail) class.
public class Invoice
{
    public int Id { get; set; }
    public decimal Amount { get; set; }
}
