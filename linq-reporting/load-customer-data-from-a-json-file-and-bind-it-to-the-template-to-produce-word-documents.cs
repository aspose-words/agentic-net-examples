using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File paths.
        string jsonPath = "customers.json";
        string templatePath = "CustomerTemplate.docx";
        string outputPath = "CustomerReport.docx";

        // Create sample JSON data.
        var sampleData = new CustomerReport
        {
            Customers = new List<Customer>
            {
                new() { Name = "Alice Johnson", Email = "alice@example.com", Phone = "555-0101" },
                new() { Name = "Bob Smith", Email = "bob@example.com", Phone = "555-0202" },
                new() { Name = "Carol Davis", Email = "carol@example.com", Phone = "555-0303" }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load JSON into the model.
        var json = File.ReadAllText(jsonPath);
        var model = JsonConvert.DeserializeObject<CustomerReport>(json)!;

        // Create the LINQ Reporting template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln();

        // Begin foreach loop over Customers.
        builder.Writeln("<<foreach [c in Customers]>>");

        // Table for each customer (header + data row).
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Email");
        builder.InsertCell();
        builder.Writeln("Phone");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[c.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[c.Email]>>");
        builder.InsertCell();
        builder.Writeln("<<[c.Phone]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using the model.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}

// Root model class.
public class CustomerReport
{
    public List<Customer> Customers { get; set; } = new();
}

// Customer data class.
public class Customer
{
    public string Name { get; set; } = "";
    public string Email { get; set; } = "";
    public string Phone { get; set; } = "";
}
