using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Invoice
{
    public int Id { get; set; }
    public decimal Amount { get; set; }
    public DateTime Date { get; set; }
}

public class ReportModel
{
    public List<Invoice> Invoices { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a foreach tag to iterate over the Invoices collection.
        builder.Writeln("<<foreach [invoice in Invoices]>>");
        builder.Writeln("Invoice ID: <<[invoice.Id]>>");
        builder.Writeln("Amount: $<<[invoice.Amount]>>");
        builder.Writeln("Date: <<[invoice.Date.ToString(\"yyyy-MM-dd\")]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "InvoiceTemplate.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Invoices = new()
            {
                new Invoice { Id = 1, Amount = 199.99m, Date = DateTime.Today.AddDays(-2) },
                new Invoice { Id = 2, Amount = 349.50m, Date = DateTime.Today.AddDays(-1) },
                new Invoice { Id = 3, Amount = 89.75m, Date = DateTime.Today }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "InvoiceReport.docx";
        doc.Save(outputPath);
    }
}
