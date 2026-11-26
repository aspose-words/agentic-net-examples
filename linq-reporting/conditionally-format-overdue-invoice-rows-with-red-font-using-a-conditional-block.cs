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
    public bool IsOverdue { get; set; }
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
            Invoices = new List<Invoice>
            {
                new Invoice { Id = 1, Date = DateTime.Today.AddDays(-10), Amount = 150.00m, IsOverdue = true },
                new Invoice { Id = 2, Date = DateTime.Today.AddDays(-5),  Amount = 200.00m, IsOverdue = false },
                new Invoice { Id = 3, Date = DateTime.Today.AddDays(-20), Amount = 350.00m, IsOverdue = true }
            }
        };

        // Create the template document.
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Invoice Report");
        builder.Writeln();

        // Begin foreach block.
        builder.Writeln("<<foreach [inv in Invoices]>>");

        // Table for each invoice row.
        Table table = builder.StartTable();

        // Id column with conditional red font.
        builder.InsertCell();
        builder.Writeln(
            @"<<if [inv.IsOverdue]>> <<textColor [""Red""]>><<[inv.Id]>> <</textColor>><</if>>" +
            @"<<if [!inv.IsOverdue]>> <<[inv.Id]>> <</if>>");

        // Date column with same conditional formatting.
        builder.InsertCell();
        builder.Writeln(
            @"<<if [inv.IsOverdue]>> <<textColor [""Red""]>><<[inv.Date]>> <</textColor>><</if>>" +
            @"<<if [!inv.IsOverdue]>> <<[inv.Date]>> <</if>>");

        // Amount column with same conditional formatting.
        builder.InsertCell();
        builder.Writeln(
            @"<<if [inv.IsOverdue]>> <<textColor [""Red""]>><<[inv.Amount]>> <</textColor>><</if>>" +
            @"<<if [!inv.IsOverdue]>> <<[inv.Amount]>> <</if>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report.
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
