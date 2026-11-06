using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Financial Report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Description: <<[item.Description]>>");
        builder.Writeln("Original Amount: <<[item.Amount]>>");
        // Use a property on the data model to perform rounding.
        builder.Writeln("Rounded Amount: <<[item.RoundedAmount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a file.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items =
            {
                new Item { Description = "Consulting Services", Amount = 1234.567m },
                new Item { Description = "Software License", Amount = 987.654m },
                new Item { Description = "Maintenance", Amount = 250.0m }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model representing a financial entry.
public class Item
{
    public string Description { get; set; } = "";
    public decimal Amount { get; set; }

    // Rounded amount to two decimal places.
    public decimal RoundedAmount => decimal.Round(Amount, 2);
}
