using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Invoice
{
    public int Id { get; set; }
    public DateTime DueDate { get; set; }
    public decimal Amount { get; set; }
    public bool IsOverdue => DueDate.Date < DateTime.Today;
}

public class ReportModel
{
    public List<Invoice> Invoices { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        var model = new ReportModel
        {
            Invoices = new List<Invoice>
            {
                new Invoice { Id = 1, DueDate = DateTime.Today.AddDays(-10), Amount = 150.00m },
                new Invoice { Id = 2, DueDate = DateTime.Today.AddDays(5), Amount = 250.00m },
                new Invoice { Id = 3, DueDate = DateTime.Today.AddDays(-2), Amount = 99.99m },
                new Invoice { Id = 4, DueDate = DateTime.Today.AddDays(15), Amount = 500.00m }
            }
        };

        // Create template.
        const string templatePath = "InvoiceReportTemplate.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Invoice Report");
        builder.Writeln();

        // Begin foreach block – each iteration will generate its own table.
        builder.Writeln("<<foreach [invoice in Invoices]>>");

        // Start a table for the current invoice.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell(); builder.Writeln("Id");
        builder.InsertCell(); builder.Writeln("Due Date");
        builder.InsertCell(); builder.Writeln("Amount");
        builder.EndRow();

        // Data row with conditional red text for overdue invoices.
        builder.InsertCell();
        builder.Writeln(
            "<<if [invoice.IsOverdue]>>" +
            "<<textColor [\"Red\"]>><<[invoice.Id]>> <</textColor>><</if>>" +
            "<<if [!invoice.IsOverdue]>>" +
            "<<[invoice.Id]>>" +
            "<</if>>");

        builder.InsertCell();
        builder.Writeln(
            "<<if [invoice.IsOverdue]>>" +
            "<<textColor [\"Red\"]>><<[invoice.DueDate]>> <</textColor>><</if>>" +
            "<<if [!invoice.IsOverdue]>>" +
            "<<[invoice.DueDate]>>" +
            "<</if>>");

        builder.InsertCell();
        builder.Writeln(
            "<<if [invoice.IsOverdue]>>" +
            "<<textColor [\"Red\"]>><<[invoice.Amount]>> <</textColor>><</if>>" +
            "<<if [!invoice.IsOverdue]>>" +
            "<<[invoice.Amount]>>" +
            "<</if>>");

        builder.EndRow();

        // End the table for this invoice.
        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        const string outputPath = "InvoiceReport.docx";
        reportDoc.Save(outputPath);
    }
}
