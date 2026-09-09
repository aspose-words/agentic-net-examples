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
                new Customer { Name = "John Doe", Address = "123 Main St, Anytown" },
                new Customer { Name = "Jane Smith", Address = "456 Oak Ave, Othertown" }
            }
        };

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Build the report using LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // No special options required for this simple example.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }

    // Creates a simple Word template containing LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title.
        builder.Writeln("Customer Report");
        builder.Writeln("----------------");

        // Begin a foreach loop over the Customers collection.
        builder.Writeln("<<foreach [c in Customers]>>");

        // Insert fields for each customer's data.
        builder.Writeln("Name   : <<[c.Name]>>");
        builder.Writeln("Address: <<[c.Address]>>");
        builder.Writeln(""); // Empty line between records.

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Wrapper class that will be passed as the root data source.
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

// Simple data model representing a customer.
public class Customer
{
    public string Name { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
}
