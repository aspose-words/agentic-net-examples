using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

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
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // Create the template document with LINQ Reporting tags
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("<<foreach [c in Customers]>>");
        // Use an expression that converts the name to upper case directly
        builder.Writeln("<<[c.Name.ToUpper()]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // Load the template (simulating a separate load step)
        Document doc = new Document(templatePath);

        // Prepare sample data
        ReportModel model = new()
        {
            Customers = new()
            {
                new Customer { Name = "Alice" },
                new Customer { Name = "Bob" },
                new Customer { Name = "Charlie" }
            }
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
