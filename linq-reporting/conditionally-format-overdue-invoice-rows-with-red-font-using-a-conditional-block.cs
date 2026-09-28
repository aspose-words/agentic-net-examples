using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Invoice
{
    public int Id { get; set; }
    public DateTime Date { get; set; }
    public decimal Amount { get; set; }
    public DateTime DueDate { get; set; }

    // Overdue if current date is later than the due date.
    public bool IsOverdue => DateTime.Now > DueDate;
}

public class ReportModel
{
    public List<Invoice> Invoices { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel
        {
            Invoices = new()
            {
                new Invoice { Id = 1, Date = DateTime.Today.AddDays(-30), Amount = 150.00m, DueDate = DateTime.Today.AddDays(-10) },
                new Invoice { Id = 2, Date = DateTime.Today.AddDays(-20), Amount = 250.00m, DueDate = DateTime.Today.AddDays(5) },
                new Invoice { Id = 3, Date = DateTime.Today.AddDays(-15), Amount = 300.00m, DueDate = DateTime.Today.AddDays(-1) },
                new Invoice { Id = 4, Date = DateTime.Today.AddDays(-5),  Amount = 120.00m, DueDate = DateTime.Today.AddDays(10) }
            }
        };

        // Create the LINQ Reporting template.
        const string templatePath = "InvoiceTemplate.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Invoice Report");
        builder.Writeln();

        // Begin foreach block.
        builder.Writeln("<<foreach [inv in Invoices]>>");

        // Table for each invoice row.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell(); builder.Writeln("ID");
        builder.InsertCell(); builder.Writeln("Date");
        builder.InsertCell(); builder.Writeln("Amount");
        builder.InsertCell(); builder.Writeln("Due Date");
        builder.EndRow();

        // Data row (template).
        builder.InsertCell();
        builder.Writeln(
            "<<if [inv.IsOverdue]>>" +
            "<<textColor [\"Red\"]>><<[inv.Id]>> <</textColor>><</if>>" +
            "<<if [!inv.IsOverdue]>>" +
            "<<[inv.Id]>>" +
            "<</if>>");

        builder.InsertCell(); builder.Writeln("<<[inv.Date]>>");
        builder.InsertCell(); builder.Writeln("<<[inv.Amount]>>");
        builder.InsertCell(); builder.Writeln("<<[inv.DueDate]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "InvoiceReport.docx";
        reportDoc.Save(outputPath);
    }
}
