using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample data.
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order
                {
                    ClientName = "Acme Corp",
                    Services = new List<Service>
                    {
                        new Service { Name = "Consulting" },
                        new Service { Name = "Support" }
                    }
                },
                new Order
                {
                    ClientName = "Globex Inc",
                    Services = new List<Service>
                    {
                        new Service { Name = "Implementation" },
                        new Service { Name = "Training" },
                        new Service { Name = "Maintenance" }
                    }
                }
            }
        };

        // Build the template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin a foreach over the orders collection.
        builder.Writeln("<<foreach [order in Orders]>>");
        // Output the client name.
        builder.Writeln("<<[order.ClientName]>>");
        // Numbered list of services – restart numbering for each order.
        builder.Writeln("1. <<restartNum>><<foreach [service in order.Services]>> <<[service.Name]>> <</foreach>>");
        // End the orders foreach.
        builder.Writeln("<</foreach>>");

        // Save the template (optional, demonstrates the lifecycle rule).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template (demonstrates the load rule).
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Root data model.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

// Order with a collection of services.
public class Order
{
    public string ClientName { get; set; } = string.Empty;
    public List<Service> Services { get; set; } = new();
}

// Simple service item.
public class Service
{
    public string Name { get; set; } = string.Empty;
}
