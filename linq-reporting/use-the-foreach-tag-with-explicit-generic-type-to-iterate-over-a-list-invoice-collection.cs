using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Invoice
{
    public int Id { get; set; }
    public decimal Amount { get; set; }
    public DateTime Date { get; set; } = DateTime.Now;
}

public class ReportModel
{
    public List<Invoice> Invoices { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Invoices = new()
            {
                new Invoice { Id = 1001, Amount = 250.75m },
                new Invoice { Id = 1002, Amount = 480.00m },
                new Invoice { Id = 1003, Amount = 125.50m }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Invoice Report");
        // Correct foreach syntax: no explicit generic type inside the tag.
        builder.Writeln("<<foreach [invoice in model.Invoices]>>");
        builder.Writeln("Invoice #: <<[invoice.Id]>>");
        builder.Writeln("Amount: $<<[invoice.Amount]>>");
        builder.Writeln("Date: <<[invoice.Date.ToString(\"yyyy-MM-dd\")]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
