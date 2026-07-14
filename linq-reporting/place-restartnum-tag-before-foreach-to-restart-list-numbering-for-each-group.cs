using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list style that will be used for the service items.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list; // Apply the list to subsequent paragraphs.

        // LINQ Reporting template.
        // Outer loop over orders.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Client: <<[order.ClientName]>>");

        // Restart numbering for each order's service list.
        // The <<restartNum>> tag must be placed before the <<foreach>> tag in the same numbered paragraph.
        builder.Writeln("<<restartNum>><<foreach [service in order.Services]>>");
        builder.Writeln("<<[service.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Prepare sample data.
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

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string ClientName { get; set; } = "";
    public List<Service> Services { get; set; } = new();
}

public class Service
{
    public string Name { get; set; } = "";
}
