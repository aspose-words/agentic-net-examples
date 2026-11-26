using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "John Doe";
    public List<Service> Services { get; set; } = new()
    {
        new Service { Name = "Consulting" },
        new Service { Name = "Support" },
        new Service { Name = "Implementation" }
    };
}

public class Service
{
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Services:");
        builder.Writeln("<<foreach [service in order.Services]>>");
        builder.Writeln("- <<[service.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        Order order = new();

        // Disable reflection optimization within this scoped block.
        bool previousOptimization = ReportingEngine.UseReflectionOptimization;
        ReportingEngine.UseReflectionOptimization = false;
        try
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");
        }
        finally
        {
            // Restore the original setting.
            ReportingEngine.UseReflectionOptimization = previousOptimization;
        }

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
