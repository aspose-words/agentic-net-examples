using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample CSV file with customer data.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "customers.csv");
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Email,Address",
            "Alice Johnson,alice@example.com,123 Maple St.",
            "Bob Smith,bob@example.com,456 Oak Ave.",
            "Carol Davis,carol@example.com,789 Pine Rd."
        });

        // 2. Create a Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        CreateTemplate(templatePath);

        // 3. Load customers from the CSV file.
        List<Customer> customers = LoadCustomersFromCsv(csvPath);

        // 4. Generate a personalized report for each customer.
        foreach (Customer customer in customers)
        {
            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare the root model for the report.
            ReportModel model = new ReportModel { Customer = customer };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the personalized report.
            string reportPath = Path.Combine(outputDir, $"{SanitizeFileName(customer.Name)}_Report.docx");
            doc.Save(reportPath);
        }
    }

    // Creates a simple Word template containing LINQ Reporting tags.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Personalized Report");
        builder.Writeln("-------------------");
        builder.Writeln("Name: <<[model.Customer.Name]>>");
        builder.Writeln("Email: <<[model.Customer.Email]>>");
        builder.Writeln("Address: <<[model.Customer.Address]>>");

        doc.Save(path);
    }

    // Parses the CSV file into a list of Customer objects.
    private static List<Customer> LoadCustomersFromCsv(string csvPath)
    {
        var customers = new List<Customer>();
        string[] lines = File.ReadAllLines(csvPath);

        // Assume the first line contains headers.
        for (int i = 1; i < lines.Length; i++)
        {
            string line = lines[i];
            if (string.IsNullOrWhiteSpace(line))
                continue;

            string[] parts = line.Split(',');
            if (parts.Length < 3)
                continue;

            customers.Add(new Customer
            {
                Name = parts[0].Trim(),
                Email = parts[1].Trim(),
                Address = parts[2].Trim()
            });
        }

        return customers;
    }

    // Removes invalid characters from a file name.
    private static string SanitizeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}

// Root model used by the LINQ Reporting engine.
public class ReportModel
{
    public Customer Customer { get; set; } = new Customer();
}

// Simple data class representing a customer.
public class Customer
{
    public string Name { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
}
