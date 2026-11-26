using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

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
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data
        var model = new ReportModel
        {
            Orders = new()
            {
                new Order { CustomerName = "Alice",   Total = 120.50m },
                new Order { CustomerName = "Bob",     Total = 75.00m },
                new Order { CustomerName = "Charlie", Total = 200.00m }
            }
        };

        // -----------------------------------------------------------------
        // Create the template document programmatically
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln();

        // Header table (static, not repeated)
        Table headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Customer");
        builder.InsertCell();
        builder.Writeln("Total");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln(); // blank line before data rows

        // Data rows – generated with a foreach tag
        builder.Writeln("<<foreach [order in Orders]>>");

        Table dataTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[order.CustomerName]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.Total]>>");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        builder.Writeln(); // blank line before summary

        // Summary using a LINQ expression inside a tag
        builder.Writeln("Grand Total: <<[Orders.Sum(o => o.Total)]>>");

        // Save the template
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model"
        engine.BuildReport(reportDoc, model, "model");

        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
