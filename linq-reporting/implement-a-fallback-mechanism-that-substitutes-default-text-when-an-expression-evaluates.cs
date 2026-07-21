using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Template: iterate over the collection "customers".
        builder.Writeln("<<foreach [c in customers]>>");
        // Use the null‑coalescing operator to provide a fallback when Name is null.
        builder.Writeln("Customer: <<[c.Name ?? \"(no name)\"]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        var model = new ReportModel
        {
            Customers = new List<Customer>
            {
                new Customer { Name = "Alice" },
                new Customer { Name = null } // This entry will trigger the fallback text.
            }
        };

        // Configure the reporting engine.
        var engine = new ReportingEngine
        {
            // Treat missing members as null (not strictly required here but safe).
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report using the root object name "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Root data model referenced by the template.
public class ReportModel
{
    // The collection name must match the tag used in the template ("customers").
    public List<Customer> Customers { get; set; } = new();
}

// Simple data entity with a nullable property.
public class Customer
{
    public string? Name { get; set; }
}
