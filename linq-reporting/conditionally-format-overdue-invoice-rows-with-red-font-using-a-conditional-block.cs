using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Invoice
{
    public int Id { get; set; }
    public DateTime IssueDate { get; set; }
    public DateTime DueDate { get; set; }
    public decimal Amount { get; set; }

    // True when the invoice is overdue.
    public bool IsOverdue => DueDate < DateTime.Today;
}

public class ReportModel
{
    public List<Invoice> Invoices { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Prepare sample data.
        var model = new ReportModel
        {
            Invoices = new List<Invoice>
            {
                new Invoice { Id = 1, IssueDate = DateTime.Today.AddDays(-30), DueDate = DateTime.Today.AddDays(-5), Amount = 1500.00m },
                new Invoice { Id = 2, IssueDate = DateTime.Today.AddDays(-20), DueDate = DateTime.Today.AddDays(10), Amount =  750.00m },
                new Invoice { Id = 3, IssueDate = DateTime.Today.AddDays(-10), DueDate = DateTime.Today.AddDays(-1), Amount =  300.00m }
            }
        };

        // 2. Create the LINQ Reporting template programmatically.
        const string templatePath = "InvoiceTemplate.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin the foreach block over the Invoices collection.
        builder.Writeln("<<foreach [inv in model.Invoices]>>");

        // Create a table with a header row.
        var table = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Id");
        builder.InsertCell(); builder.Writeln("Issue Date");
        builder.InsertCell(); builder.Writeln("Due Date");
        builder.InsertCell(); builder.Writeln("Amount");
        builder.EndRow();

        // Data row – each cell contains a field expression.
        builder.InsertCell(); builder.Writeln("<<[inv.Id]>>");
        builder.InsertCell(); builder.Writeln("<<[inv.IssueDate]>>");
        builder.InsertCell(); builder.Writeln("<<[inv.DueDate]>>");

        // Conditional formatting for the Amount column.
        builder.InsertCell();
        builder.Writeln(
            "<<if [inv.IsOverdue]>>" +
            "<<textColor [\"Red\"]>><<[inv.Amount]>> <</textColor>><</if>>" +
            "<<if [!inv.IsOverdue]>>" +
            "<<[inv.Amount]>>" +
            "<</if>>");

        builder.EndRow();
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template and build the report.
        var loadedTemplate = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(loadedTemplate, model, "model");

        // 4. Save the generated report.
        const string reportPath = "InvoiceReport.docx";
        loadedTemplate.Save(reportPath);
    }
}
