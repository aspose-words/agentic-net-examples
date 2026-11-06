using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Customer
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer List:");
        builder.Writeln("<<foreach [c in customers]>>");
        builder.Writeln("Name: <<[c.Name]>>, Age: <<[c.Age]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Prepare a collection of customers.
        var customers = new List<Customer>
        {
            new() { Name = "Alice", Age = 30 },
            new() { Name = "Bob", Age = 25 },
            new() { Name = "Charlie", Age = 28 }
        };

        // Build the report, specifying the collection name "customers".
        var engine = new ReportingEngine();
        engine.BuildReport(doc, customers, "customers");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
