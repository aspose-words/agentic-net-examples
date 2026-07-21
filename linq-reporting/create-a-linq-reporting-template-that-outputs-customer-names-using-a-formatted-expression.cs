using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Customer
{
    // Initialize to avoid nullable warnings.
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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple LINQ Reporting template.
        // The template will iterate over the Customers collection and output each name.
        builder.Writeln("Customer List:");
        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln(" - <<[c.Name]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel();
        model.Customers.Add(new Customer { Name = "Alice Johnson" });
        model.Customers.Add(new Customer { Name = "Bob Smith" });
        model.Customers.Add(new Customer { Name = "Charlie Davis" });

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The root object name must match the name used in the template tags.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("LINQReporting_Output.docx");
    }
}
