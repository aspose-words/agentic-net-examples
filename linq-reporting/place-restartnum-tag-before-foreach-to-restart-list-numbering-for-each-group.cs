using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Service
{
    public string Name { get; set; } = "";
}

public class Order
{
    public string ClientName { get; set; } = "";
    public List<Service> Services { get; set; } = new();
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
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
                        new Service { Name = "Implementation" },
                        new Service { Name = "Support" }
                    }
                },
                new Order
                {
                    ClientName = "Globex Inc",
                    Services = new List<Service>
                    {
                        new Service { Name = "Analysis" },
                        new Service { Name = "Design" }
                    }
                }
            }
        };

        // Create the template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Outer foreach over Orders.
        builder.Writeln("<<foreach [order in Orders]>>");
        // Write the client name.
        builder.Writeln("Client: <<[order.ClientName]>>");

        // Start a numbered list for the services.
        builder.ListFormat.List = template.Lists.Add(ListTemplate.NumberDefault);

        // Restart numbering for each order.
        builder.Writeln("<<restartNum>>");

        // Inner foreach over Services.
        builder.Writeln("<<foreach [service in order.Services]>>");
        builder.Writeln("<<[service.Name]>>");
        builder.Writeln("<</foreach>>");

        // End the list for this order.
        builder.ListFormat.RemoveNumbers();

        // Close the outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template (simulating a separate load step).
        var loadedTemplate = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        loadedTemplate.Save("Report.docx");
    }
}
