using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for potential encoding needs.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        ReportModel model = new()
        {
            Customers = new()
            {
                new Customer
                {
                    Name = "Acme Corp",
                    Invoices = new()
                    {
                        new Invoice { Number = "INV-001", Date = new DateTime(2023, 1, 15), Amount = 1234.56m },
                        new Invoice { Number = "INV-002", Date = new DateTime(2023, 2, 20), Amount = 789.00m }
                    }
                },
                new Customer
                {
                    Name = "Globex Inc",
                    Invoices = new()
                    {
                        new Invoice { Number = "INV-101", Date = new DateTime(2023, 3, 5), Amount = 2500.00m }
                    }
                }
            }
        };

        // Create the LINQ Reporting template programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("<<foreach [customer in Customers]>>");
        builder.Writeln("Customer: <<[customer.Name]>>");
        builder.Writeln("Invoices:");
        builder.Writeln("<<foreach [invoice in customer.Invoices]>>");
        builder.Writeln("- Invoice #: <<[invoice.Number]>>, Date: <<[invoice.Date]>>, Amount: <<[invoice.Amount]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

public class Customer
{
    public string Name { get; set; } = "";
    public List<Invoice> Invoices { get; set; } = new();
}

public class Invoice
{
    public string Number { get; set; } = "";
    public DateTime Date { get; set; }
    public decimal Amount { get; set; }
}
