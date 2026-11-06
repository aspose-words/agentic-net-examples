using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Customer
{
    public int Id { get; set; } = 0;
    public string Name { get; set; } = "";
    public string Email { get; set; } = "";
    public string Address { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Reports");
        Directory.CreateDirectory(outputDir);

        // Create sample CSV file with customer data.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "customers.csv");
        var csvLines = new[]
        {
            "Id,Name,Email,Address",
            "1,John Doe,john.doe@example.com,123 Main St",
            "2,Jane Smith,jane.smith@example.com,456 Oak Ave",
            "3,Bob Johnson,bob.johnson@example.com,789 Pine Rd"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // Load customers from CSV into a list of Customer objects.
        var customers = new List<Customer>();
        using (var reader = new StreamReader(csvPath))
        {
            // Read header.
            string? headerLine = reader.ReadLine();
            if (headerLine == null) return;

            // Read data rows.
            while (!reader.EndOfStream)
            {
                string? line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) continue;

                string[] parts = line.Split(',');
                if (parts.Length < 4) continue;

                customers.Add(new Customer
                {
                    Id = int.Parse(parts[0], CultureInfo.InvariantCulture),
                    Name = parts[1],
                    Email = parts[2],
                    Address = parts[3]
                });
            }
        }

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "CustomerTemplate.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[customer.Name]>>");
        builder.Writeln("Email: <<[customer.Email]>>");
        builder.Writeln("Address: <<[customer.Address]>>");

        // Save the template for later reuse.
        templateDoc.Save(templatePath);

        // Generate a personalized report for each customer.
        foreach (var customer in customers)
        {
            // Load the template fresh for each report.
            var reportDoc = new Document(templatePath);

            // Build the report using the customer as the root object.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, customer, "customer");

            // Save the generated report.
            string reportPath = Path.Combine(outputDir, $"Report_{customer.Id}.docx");
            reportDoc.Save(reportPath);
        }
    }
}
