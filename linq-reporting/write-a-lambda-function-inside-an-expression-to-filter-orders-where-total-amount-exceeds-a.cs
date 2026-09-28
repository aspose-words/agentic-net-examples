using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "";
    public decimal Total { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
    public decimal Threshold { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Paths for template and output documents.
        string templatePath = "template.docx";
        string outputPath = "report.docx";

        // Create a template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("Orders with Total > Threshold");
        builder.Writeln();

        // Foreach loop with a lambda expression to filter orders.
        builder.Writeln("<<foreach [order in Orders.Where(o => o.Total > Threshold)]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>, Total: <<[order.Total]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Threshold = 150m,
            Orders = new()
            {
                new Order { CustomerName = "Alice", Total = 120m },
                new Order { CustomerName = "Bob",   Total = 200m },
                new Order { CustomerName = "Carol", Total = 300m }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
