using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings Aspose.Words might need.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title.
        builder.Writeln("Orders Report");
        builder.Writeln();

        // Outer foreach iterates over the collection of orders.
        builder.Writeln("<<foreach [order in Orders]>>");

        // Display the client name.
        builder.Writeln("Client: <<[order.ClientName]>>");

        // Numbered list of services for the current order.
        // The <<restartNum>> tag restarts numbering for each order.
        builder.Writeln("1. <<restartNum>><<foreach [service in order.Services]>> <<[service.Name]>> <</foreach>>");

        // End of the outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
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

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this example.
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, non‑nullable properties are initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string ClientName { get; set; } = string.Empty;
    public List<Service> Services { get; set; } = new();
}

public class Service
{
    public string Name { get; set; } = string.Empty;
}
