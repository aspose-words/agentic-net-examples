using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create the template document once.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Sample data for two isolated requests.
        var order1 = new Order
        {
            CustomerName = "Alice Johnson",
            Services = new List<Service>
            {
                new Service { Name = "Consulting" },
                new Service { Name = "Support" }
            }
        };

        var order2 = new Order
        {
            CustomerName = "Bob Smith",
            Services = new List<Service>
            {
                new Service { Name = "Installation" },
                new Service { Name = "Training" },
                new Service { Name = "Maintenance" }
            }
        };

        // Process each request with its own ReportingEngine instance.
        ProcessReport(order1, templatePath, "Report1.docx");
        ProcessReport(order2, templatePath, "Report2.docx");
    }

    // Creates a simple Word template containing LINQ Reporting tags.
    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Services:");
        builder.Writeln("<<foreach [service in order.Services]>>- <<[service.Name]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }

    // Generates a report for a single request using an isolated ReportingEngine.
    private static void ProcessReport(Order data, string templatePath, string outputPath)
    {
        // Load the template fresh for this request.
        var doc = new Document(templatePath);

        // Each request gets its own ReportingEngine instance.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Build the report; the root object name must match the tag prefix.
        bool success = engine.BuildReport(doc, data, "order");

        // Optionally, handle build failures (e.g., when InlineErrorMessages is used).
        if (!success)
        {
            Console.WriteLine($"Report generation failed for {outputPath}");
        }

        doc.Save(outputPath);
    }
}

// Public data model aligned with the template tags.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Service> Services { get; set; } = new();
}

public class Service
{
    public string Name { get; set; } = string.Empty;
}
