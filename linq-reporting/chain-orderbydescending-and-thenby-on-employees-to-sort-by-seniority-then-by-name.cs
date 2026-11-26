using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingSortingExample
{
    // Data model classes
    public class Company
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Employee
    {
        public string Name { get; set; } = "";
        public int Seniority { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required by Aspose.Words for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample employee data
            List<Employee> employees = new()
            {
                new() { Name = "Alice", Seniority = 5 },
                new() { Name = "Bob", Seniority = 7 },
                new() { Name = "Charlie", Seniority = 5 },
                new() { Name = "David", Seniority = 3 }
            };

            // Chain OrderByDescending and ThenBy to sort by seniority descending, then by name ascending
            List<Employee> sortedEmployees = employees
                .OrderByDescending(e => e.Seniority)
                .ThenBy(e => e.Name)
                .ToList();

            // Create the root data object
            Company model = new() { Employees = sortedEmployees };

            // Define file paths
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string reportPath = Path.Combine(outputDir, "Report.docx");

            // -----------------------------------------------------------------
            // Step 1: Create the LINQ Reporting template programmatically
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Employee Report");
            builder.Writeln("<<foreach [e in Employees]>>");
            builder.Writeln("<<[e.Name]>> - Seniority: <<[e.Seniority]>>");
            builder.Writeln("<</foreach>>");

            // Save the template
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            reportDoc.Save(reportPath);
        }
    }
}
