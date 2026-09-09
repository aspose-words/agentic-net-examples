using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                new Order { CustomerName = "Alice", Total = 120.50m },
                new Order { CustomerName = "Bob",   Total =  85.75m },
                new Order { CustomerName = "Carol", Total = 210.00m }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Order Report");
        builder.Writeln();

        // List each order.
        builder.Writeln("<<foreach [order in model.Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>  -  Total: $<<[order.Total]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Summary section using LINQ aggregation inside a tag expression.
        builder.Writeln("Summary:");
        builder.Writeln("Number of Orders: <<[model.Orders.Count]>>");
        builder.Writeln("Grand Total: $<<[model.Orders.Sum(o => o.Total)]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options

        // Build the report using the model as the root data source named "model".
        bool success = engine.BuildReport(doc, model, "model");

        // Optionally check the success flag (relevant only with InlineErrorMessages).
        if (!success)
        {
            Console.WriteLine("Report generation encountered errors.");
        }

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
