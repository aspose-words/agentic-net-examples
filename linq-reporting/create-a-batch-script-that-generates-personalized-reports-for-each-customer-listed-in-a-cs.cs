using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(baseDir, "CustomerTemplate.docx");
        string csvPath = Path.Combine(baseDir, "Customers.csv");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a simple Word template with LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Email: <<[model.Email]>>");
        builder.Writeln("Address: <<[model.Address]>>");
        templateDoc.Save(templatePath);

        // 2. Create a sample CSV file with customer data.
        var csvLines = new[]
        {
            "Name,Email,Address",
            "Alice Smith,alice@example.com,123 Maple St.",
            "Bob Johnson,bob@example.com,456 Oak Ave.",
            "Carol Davis,carol@example.com,789 Pine Rd."
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // 3. Load CSV data into a list of Customer objects.
        var customers = LoadCustomersFromCsv(csvPath);

        // 4. Generate a personalized report for each customer.
        foreach (var customer in customers)
        {
            // Load a fresh copy of the template for each report.
            var reportDoc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, customer, "model");

            // Save the generated report.
            string safeName = MakeFileNameSafe(customer.Name);
            string reportPath = Path.Combine(outputDir, $"{safeName}_Report.docx");
            reportDoc.Save(reportPath);
        }
    }

    // Loads customers from a CSV file into a list of Customer objects.
    private static List<Customer> LoadCustomersFromCsv(string csvFilePath)
    {
        var customers = new List<Customer>();
        using var reader = new StreamReader(csvFilePath);
        // Read header line.
        string? headerLine = reader.ReadLine();
        if (headerLine == null) return customers;

        while (!reader.EndOfStream)
        {
            string? line = reader.ReadLine();
            if (string.IsNullOrWhiteSpace(line)) continue;

            var parts = line.Split(',');
            if (parts.Length < 3) continue;

            var customer = new Customer
            {
                Name = parts[0].Trim(),
                Email = parts[1].Trim(),
                Address = parts[2].Trim()
            };
            customers.Add(customer);
        }

        return customers;
    }

    // Replaces invalid filename characters.
    private static string MakeFileNameSafe(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}

// Public data model for a customer.
public class Customer
{
    public string Name { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
}
