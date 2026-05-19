using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template with LINQ Reporting tags.
        var templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template document.
        var doc = new Document(templatePath);

        // Prepare the data model. Intentionally omit the MissingProperty to trigger a warning.
        var model = new ReportModel
        {
            CustomerName = "Acme Corp",
            HasOrders = true,
            Orders = new List<Order>
            {
                new Order { Name = "Widget" },
                new Order { Name = "Gadget" }
            }
        };

        // Configure the reporting engine to inline error messages.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;
        engine.MissingMemberMessage = "Missing";

        // Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "output.docx";
        doc.Save(outputPath);

        // Output simple status information.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Report saved to: {outputPath}");
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Basic data insertion.
        builder.Writeln("Report for <<[model.CustomerName]>>");

        // Optional section that may be omitted.
        builder.Writeln("<<if [model.HasOrders]>>");
        builder.Writeln("Orders:");
        builder.Writeln("<<foreach [order in model.Orders]>>");
        builder.Writeln("- <<[order.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</if>>");

        // Reference to a missing member to generate an inline error.
        builder.Writeln("Missing field: <<[model.MissingProperty]>>");

        // Placeholder to capture any inline error messages.
        builder.Writeln("<<error>>");

        doc.Save(path);
    }
}

// Data model classes.
public class ReportModel
{
    public string CustomerName { get; set; } = string.Empty;
    public bool HasOrders { get; set; }
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string Name { get; set; } = string.Empty;
}
