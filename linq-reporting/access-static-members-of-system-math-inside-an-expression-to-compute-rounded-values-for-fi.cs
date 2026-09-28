using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var model = new ReportModel
        {
            Items =
            {
                new Item { Description = "Consulting Services", Amount = 1234.5678m },
                new Item { Description = "Software License", Amount = 250.5m },
                new Item { Description = "Support Fee", Amount = 89.999m }
            }
        };

        // Create the template document programmatically
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Financial Report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Description: <<[item.Description]>>");
        builder.Writeln("Original Amount: <<[item.Amount]>>");
        // Use a helper method defined in the root model to round the amount
        builder.Writeln("Rounded Amount: <<[model.Round(item.Amount, 2)]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template and build the report
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Root data model
public class ReportModel
{
    public List<Item> Items { get; set; } = new();

    // Helper method to round decimal values using System.Math
    public decimal Round(decimal value, int digits)
    {
        // Math.Round works with double, so convert to double and back to decimal
        return Convert.ToDecimal(Math.Round(Convert.ToDouble(value), digits));
    }
}

// Item model representing a financial entry
public class Item
{
    public string Description { get; set; } = string.Empty;
    public decimal Amount { get; set; }
}
