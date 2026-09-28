using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public double Total { get; set; } = 0;
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var order = new Order { Total = 620.0 };

        // Create a template document.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("Total: <<[order.Total]>>");
        builder.Writeln("Discount: <<[order.Total > 500 ? order.Total * 0.1 : 0]>>");
        doc.Save(templatePath);

        // Load the template.
        var template = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, order, "order");

        // Save the generated report.
        var outputPath = "Report.docx";
        template.Save(outputPath);
    }
}
