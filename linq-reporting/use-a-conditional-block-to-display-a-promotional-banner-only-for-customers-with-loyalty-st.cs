using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Customers = new List<Customer>
            {
                new Customer
                {
                    Name = "Alice Johnson",
                    IsLoyal = true,
                    PromoBanner = "Exclusive 20% Discount for Loyal Customers!"
                },
                new Customer
                {
                    Name = "Bob Smith",
                    IsLoyal = false,
                    PromoBanner = "Exclusive 20% Discount for Loyal Customers!"
                }
            }
        };

        // Create the LINQ Reporting template programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer Report");
        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln("Name: <<[c.Name]>>");
        // Conditional block: display promotional banner only for loyal customers.
        builder.Writeln("<<if [c.IsLoyal]>>");
        builder.Writeln("<<[c.PromoBanner]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Optionally, indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Wrapper class for the root data source.
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

// Customer data model.
public class Customer
{
    public string Name { get; set; } = "";
    public bool IsLoyal { get; set; }
    public string PromoBanner { get; set; } = "";
}
