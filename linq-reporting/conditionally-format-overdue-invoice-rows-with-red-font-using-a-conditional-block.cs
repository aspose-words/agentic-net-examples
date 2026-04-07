using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Invoice
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = "";
    public DateTime DueDate { get; set; }
    public decimal Amount { get; set; }

    // Computed property – true when the invoice is overdue.
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
        // Prepare sample data.
        var model = new ReportModel
        {
            Invoices = new List<Invoice>
            {
                new Invoice { Id = 1, CustomerName = "Acme Corp", DueDate = DateTime.Today.AddDays(-5), Amount = 1234.56m },
                new Invoice { Id = 2, CustomerName = "Beta Ltd",   DueDate = DateTime.Today.AddDays( 3), Amount = 789.00m },
                new Invoice { Id = 3, CustomerName = "Gamma LLC",  DueDate = DateTime.Today.AddDays(-1), Amount = 456.78m }
            }
        };

        // Create the template document programmatically.
        var templatePath = "InvoiceTemplate.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Invoice Report");
        builder.Writeln();

        // Begin foreach block over the invoices.
        builder.Writeln("<<foreach [inv in model.Invoices]>>");

        // Start a table for each invoice (header + data row).
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell(); builder.Writeln("ID");
        builder.InsertCell(); builder.Writeln("Customer");
        builder.InsertCell(); builder.Writeln("Due Date");
        builder.InsertCell(); builder.Writeln("Amount");
        builder.EndRow();

        // Data row – each cell uses a conditional block to apply red text when overdue.
        // ID cell
        builder.InsertCell();
        builder.Writeln(
            "<<if [inv.IsOverdue]>> <<textColor [\"Red\"]>><<[inv.Id]>> <</textColor>><</if>>" +
            "<<if [!inv.IsOverdue]>> <<[inv.Id]>> <</if>>");

        // Customer cell
        builder.InsertCell();
        builder.Writeln(
            "<<if [inv.IsOverdue]>> <<textColor [\"Red\"]>><<[inv.CustomerName]>> <</textColor>><</if>>" +
            "<<if [!inv.IsOverdue]>> <<[inv.CustomerName]>> <</if>>");

        // Due Date cell
        builder.InsertCell();
        builder.Writeln(
            "<<if [inv.IsOverdue]>> <<textColor [\"Red\"]>><<[inv.DueDate.ToString(\"yyyy-MM-dd\")]>> <</textColor>><</if>>" +
            "<<if [!inv.IsOverdue]>> <<[inv.DueDate.ToString(\"yyyy-MM-dd\")]>> <</if>>");

        // Amount cell
        builder.InsertCell();
        builder.Writeln(
            "<<if [inv.IsOverdue]>> <<textColor [\"Red\"]>><<[inv.Amount.ToString(\"C\")]>> <</textColor>><</if>>" +
            "<<if [!inv.IsOverdue]>> <<[inv.Amount.ToString(\"C\")]>> <</if>>");

        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = "InvoiceReport.docx";
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
