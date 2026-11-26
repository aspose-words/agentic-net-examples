using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Customer
{
    public string Name { get; set; } = "";
    public string Email { get; set; } = "";
    public string Address { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directories.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputDir);

        // Create sample CSV file.
        string csvPath = Path.Combine(dataDir, "customers.csv");
        File.WriteAllText(csvPath,
            "Name,Email,Address\n" +
            "Alice Johnson,alice@example.com,123 Maple Street\n" +
            "Bob Smith,bob@example.com,456 Oak Avenue\n" +
            "Carol Davis,carol@example.com,789 Pine Road",
            Encoding.UTF8);

        // Create the template document programmatically.
        string templatePath = Path.Combine(dataDir, "CustomerTemplate.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[customer.Name]>>");
        builder.Writeln("Email: <<[customer.Email]>>");
        builder.Writeln("Address: <<[customer.Address]>>");
        templateDoc.Save(templatePath);

        // Load customers from CSV.
        List<Customer> customers = LoadCustomersFromCsv(csvPath);

        // Initialize the reporting engine once.
        var engine = new ReportingEngine();

        // Generate a personalized report for each customer.
        foreach (Customer cust in customers)
        {
            // Load a fresh copy of the template for each iteration.
            var doc = new Document(templatePath);

            // Build the report using the Customer object as the root data source.
            engine.BuildReport(doc, cust, "customer");

            // Save the generated report.
            string safeName = MakeFileNameSafe(cust.Name);
            string outputPath = Path.Combine(outputDir, $"{safeName}_Report.docx");
            doc.Save(outputPath);
        }
    }

    private static List<Customer> LoadCustomersFromCsv(string csvFilePath)
    {
        var list = new List<Customer>();
        using var reader = new StreamReader(csvFilePath, Encoding.UTF8);
        string? headerLine = reader.ReadLine(); // Skip header
        if (headerLine == null)
            return list;

        while (!reader.EndOfStream)
        {
            string? line = reader.ReadLine();
            if (string.IsNullOrWhiteSpace(line))
                continue;

            // Simple CSV split; assumes no commas inside quoted fields.
            string[] parts = line.Split(',');
            if (parts.Length < 3)
                continue;

            var customer = new Customer
            {
                Name = parts[0].Trim(),
                Email = parts[1].Trim(),
                Address = parts[2].Trim()
            };
            list.Add(customer);
        }

        return list;
    }

    private static string MakeFileNameSafe(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
