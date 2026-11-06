using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare paths.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Master‑Detail Report");
        builder.Writeln();

        // Outer foreach band for customers.
        builder.Writeln("<<foreach [customer in Customers]>>");
        builder.Writeln("Customer: <<[customer.Name]>>");
        builder.Writeln();

        // Inner foreach band for invoices of the current customer.
        builder.Writeln("<<foreach [invoice in customer.Invoices]>>");
        builder.Writeln("Invoice #: <<[invoice.Number]>> | Date: <<[invoice.Date]>> | Amount: $<<[invoice.Amount]>>");
        builder.Writeln("<</foreach>>"); // End inner foreach.

        builder.Writeln("<</foreach>>"); // End outer foreach.

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Create sample data model.
        // -----------------------------------------------------------------
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
                    Name = "Globex Ltd",
                    Invoices = new()
                    {
                        new Invoice { Number = "INV-101", Date = new DateTime(2023, 3, 5), Amount = 2500.00m },
                        new Invoice { Number = "INV-102", Date = new DateTime(2023, 4, 12), Amount = 1750.75m },
                        new Invoice { Number = "INV-103", Date = new DateTime(2023, 5, 30), Amount = 320.40m }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // 4. Build the report using LINQ Reporting Engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
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
