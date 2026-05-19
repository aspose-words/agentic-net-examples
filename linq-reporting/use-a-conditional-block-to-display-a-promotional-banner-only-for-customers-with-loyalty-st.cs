using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Customer
{
    public string Name { get; set; } = string.Empty;
    public bool IsLoyal { get; set; }
}

public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Customer Report");
        builder.Writeln();

        // Begin looping over the customers collection.
        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln("Name: <<[c.Name]>>");

        // Conditional block – show promotional banner only for loyal customers.
        builder.Writeln("<<if [c.IsLoyal]>>");
        builder.Writeln("=== Promotional Banner: 20% OFF on next purchase! ===");
        builder.Writeln("<</if>>");

        // End of the foreach loop.
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        var model = new ReportModel
        {
            Customers = new List<Customer>
            {
                new Customer { Name = "Alice", IsLoyal = true },
                new Customer { Name = "Bob",   IsLoyal = false },
                new Customer { Name = "Charlie", IsLoyal = true }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("CustomerReport.docx");
    }
}
