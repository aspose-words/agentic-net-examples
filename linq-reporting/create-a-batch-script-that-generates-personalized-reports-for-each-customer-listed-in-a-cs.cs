using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace BatchReportGenerator
{
    // Model representing a customer.
    public class Customer
    {
        public string CustomerId { get; set; } = "";
        public string Name { get; set; } = "";
        public string Email { get; set; } = "";
        public decimal TotalPurchase { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for possible CSV encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare folders.
            string dataFolder = "Data";
            string outputFolder = "Reports";
            Directory.CreateDirectory(dataFolder);
            Directory.CreateDirectory(outputFolder);

            // Create a sample CSV file with customer data.
            string csvPath = Path.Combine(dataFolder, "customers.csv");
            File.WriteAllText(csvPath, "CustomerId,Name,Email,TotalPurchase\r\n" +
                                         "C001,John Doe,john.doe@example.com,1234.56\r\n" +
                                         "C002,Jane Smith,jane.smith@example.com,7890.12\r\n" +
                                         "C003,Bob Johnson,bob.johnson@example.com,345.67");

            // Load customers from CSV into a list.
            List<Customer> customers = LoadCustomersFromCsv(csvPath);

            // Create the template document programmatically.
            string templatePath = Path.Combine(dataFolder, "CustomerReportTemplate.docx");
            CreateTemplateDocument(templatePath);

            // Generate a personalized report for each customer.
            foreach (Customer customer in customers)
            {
                // Load the template.
                Document doc = new Document(templatePath);

                // Build the report using the customer as the root object.
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(doc, customer, "customer");

                // Save the generated report.
                string reportFileName = $"Report_{customer.CustomerId}.docx";
                string reportPath = Path.Combine(outputFolder, reportFileName);
                doc.Save(reportPath);
            }
        }

        // Reads the CSV file and returns a list of Customer objects.
        private static List<Customer> LoadCustomersFromCsv(string csvFilePath)
        {
            var customers = new List<Customer>();
            using (var reader = new StreamReader(csvFilePath))
            {
                // Read header line.
                string? headerLine = reader.ReadLine();
                if (headerLine == null)
                    return customers; // Empty file.

                // Read data lines.
                while (!reader.EndOfStream)
                {
                    string? line = reader.ReadLine();
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] parts = line.Split(',');
                    if (parts.Length < 4)
                        continue; // Skip malformed lines.

                    var customer = new Customer
                    {
                        CustomerId = parts[0].Trim(),
                        Name = parts[1].Trim(),
                        Email = parts[2].Trim(),
                        TotalPurchase = decimal.TryParse(parts[3].Trim(), out decimal value) ? value : 0m
                    };
                    customers.Add(customer);
                }
            }
            return customers;
        }

        // Creates a Word template with LINQ Reporting tags.
        private static void CreateTemplateDocument(string templatePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Personalized Customer Report");
            builder.Writeln("--------------------------------------------------");
            builder.Writeln("Customer ID: <<[customer.CustomerId]>>");
            builder.Writeln("Name: <<[customer.Name]>>");
            builder.Writeln("Email: <<[customer.Email]>>");
            builder.Writeln("Total Purchase: $<<[customer.TotalPurchase]>>");
            builder.Writeln("--------------------------------------------------");

            doc.Save(templatePath);
        }
    }
}
