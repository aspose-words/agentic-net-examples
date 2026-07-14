using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for the Table class

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the template document.
        string templatePath = Path.Combine(outputDir, "InvoiceTemplate.docx");
        CreateTemplate(templatePath);

        // 2. Prepare sample data.
        ReportModel model = new()
        {
            Invoices = new()
            {
                new InvoiceItem
                {
                    Description = "Consulting Services",
                    DueDate = DateTime.Today.AddDays(-5), // overdue
                    Amount = 1500.00m
                },
                new InvoiceItem
                {
                    Description = "Software License",
                    DueDate = DateTime.Today.AddDays(10), // not overdue
                    Amount = 299.99m
                },
                new InvoiceItem
                {
                    Description = "Maintenance Support",
                    DueDate = DateTime.Today.AddDays(-2), // overdue
                    Amount = 750.00m
                }
            }
        };

        // 3. Load the template and build the report.
        Document report = new(templatePath);
        ReportingEngine engine = new()
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(report, model, "model");

        // 4. Save the generated report.
        string reportPath = Path.Combine(outputDir, "InvoiceReport.docx");
        report.Save(reportPath);
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Title
        builder.Writeln("Invoice Report");
        builder.Writeln();

        // Begin foreach over the collection named Invoices.
        builder.Writeln("<<foreach [item in model.Invoices]>>");

        // Table header
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Description");
        builder.InsertCell();
        builder.Writeln("Due Date");
        builder.InsertCell();
        builder.Writeln("Amount");
        builder.EndRow();

        // Data row with conditional formatting for overdue items.
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.IsOverdue]>>" +
            "<<textColor [\"Red\"]>><<[item.Description]>> <</textColor>><</if>>" +
            "<<if [!item.IsOverdue]>>" +
            "<<[item.Description]>>" +
            "<</if>>");

        builder.InsertCell();
        builder.Writeln(
            "<<if [item.IsOverdue]>>" +
            "<<textColor [\"Red\"]>><<[item.DueDate.ToString(\"yyyy-MM-dd\")]>> <</textColor>><</if>>" +
            "<<if [!item.IsOverdue]>>" +
            "<<[item.DueDate.ToString(\"yyyy-MM-dd\")]>>" +
            "<</if>>");

        builder.InsertCell();
        builder.Writeln(
            "<<if [item.IsOverdue]>>" +
            "<<textColor [\"Red\"]>><<[item.Amount]>> <</textColor>><</if>>" +
            "<<if [!item.IsOverdue]>>" +
            "<<[item.Amount]>>" +
            "<</if>>");

        builder.EndRow();
        builder.EndTable();

        // Close foreach block.
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}

// Root data model.
public class ReportModel
{
    public List<InvoiceItem> Invoices { get; set; } = new();
}

// Individual invoice item.
public class InvoiceItem
{
    public string Description { get; set; } = string.Empty;
    public DateTime DueDate { get; set; }
    public decimal Amount { get; set; }

    // Computed property used in the template to decide formatting.
    public bool IsOverdue => DueDate < DateTime.Today;
}
