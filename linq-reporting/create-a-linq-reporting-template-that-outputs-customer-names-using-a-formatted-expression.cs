using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Customer
{
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Customers = new List<Customer>
            {
                new Customer { Name = "Alice Johnson" },
                new Customer { Name = "Bob Smith" },
                new Customer { Name = "Charlie Davis" }
            }
        };

        // Create a template document programmatically.
        var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer Names:");
        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln(" - <<[c.Name]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template (optional, demonstrates load step).
        var loadedDoc = new Document(templatePath);

        // Build the report using LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedDoc, model, "model");

        // Save the generated report.
        var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        loadedDoc.Save(reportPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {reportPath}");
    }
}
