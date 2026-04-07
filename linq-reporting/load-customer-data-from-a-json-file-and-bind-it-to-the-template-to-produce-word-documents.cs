using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Customer
{
    // Initialize properties to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
}

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Register code page provider required by Aspose.Words for some encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths relative to the current working directory.
        string jsonPath = Path.Combine(Environment.CurrentDirectory, "Customers.json");
        string templatePath = Path.Combine(Environment.CurrentDirectory, "CustomerTemplate.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomerReport.docx");

        // -----------------------------------------------------------------
        // Step 1: Create sample customer data and write it to a JSON file.
        // -----------------------------------------------------------------
        List<Customer> customers = new()
        {
            new Customer { Name = "Alice Johnson", Address = "123 Maple Street, Springfield", Email = "alice@example.com" },
            new Customer { Name = "Bob Smith", Address = "456 Oak Avenue, Metropolis", Email = "bob@example.com" },
            new Customer { Name = "Carol Davis", Address = "789 Pine Road, Gotham", Email = "carol@example.com" }
        };

        // Serialize the list to JSON with indentation for readability.
        string jsonContent = JsonConvert.SerializeObject(customers, Formatting.Indented);
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // ---------------------------------------------------------------
        // Step 2: Create a Word template containing LINQ Reporting tags.
        // ---------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Customer Report");
        builder.Writeln();

        // Begin foreach loop over the JSON array (root name will be "customers").
        builder.Writeln("<<foreach [c in customers]>>");
        builder.Writeln("Name   : <<[c.Name]>>");
        builder.Writeln("Address: <<[c.Address]>>");
        builder.Writeln("Email  : <<[c.Email]>>");
        builder.Writeln(); // Blank line between records.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // Step 3: Load the template and bind the JSON data source.
        // ---------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Create a JsonDataSource from the JSON file.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report. The root object name must match the tag ("customers").
        bool success = engine.BuildReport(reportDoc, jsonDataSource, "customers");

        // Optional: you could check the success flag if InlineErrorMessages were enabled.
        // For this example we proceed regardless.

        // ---------------------------------------------------------------
        // Step 4: Save the generated report.
        // ---------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
