using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Customer
{
    public string Name { get; set; } = "";
    public string Address { get; set; } = "";
    public string Email { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file with customer data.
        // -----------------------------------------------------------------
        string csvPath = "customers.csv";
        string[] csvLines =
        {
            "Name,Address,Email",
            "Alice Johnson,123 Maple St.,alice@example.com",
            "Bob Smith,456 Oak Ave.,bob@example.com",
            "Carol Lee,789 Pine Rd.,carol@example.com"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Load CSV data into a list of Customer objects.
        // -----------------------------------------------------------------
        var customers = new List<Customer>();
        using (var reader = new StreamReader(csvPath))
        {
            // Skip header.
            string? header = reader.ReadLine();
            while (!reader.EndOfStream)
            {
                string? line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) continue;

                // Simple CSV split (no quoted commas handling needed for this sample).
                string[] parts = line.Split(',');
                if (parts.Length >= 3)
                {
                    customers.Add(new Customer
                    {
                        Name = parts[0].Trim(),
                        Address = parts[1].Trim(),
                        Email = parts[2].Trim()
                    });
                }
            }
        }

        // -----------------------------------------------------------------
        // 3. Create a Word template with LINQ Reporting tags.
        // -----------------------------------------------------------------
        string templatePath = "CustomerTemplate.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Personalized Report");
        builder.Writeln("--------------------");
        builder.Writeln("Name   : <<[Customer.Name]>>");
        builder.Writeln("Address: <<[Customer.Address]>>");
        builder.Writeln("Email  : <<[Customer.Email]>>");
        builder.Writeln(); // blank line between reports

        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Generate a separate report file for each customer.
        // -----------------------------------------------------------------
        foreach (var customer in customers)
        {
            // Load the template for each iteration to start from a clean document.
            var reportDoc = new Document(templatePath);

            // Build the report using the current customer as the data source.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, customer, "Customer");

            // Create a safe file name.
            string safeName = MakeFileNameSafe(customer.Name);
            string outputPath = $"Report_{safeName}.docx";

            reportDoc.Save(outputPath);
            Console.WriteLine($"Generated report: {outputPath}");
        }
    }

    // Helper to replace invalid filename characters.
    private static string MakeFileNameSafe(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
