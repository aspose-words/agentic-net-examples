using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // External types that will be registered with the ReportingEngine.
    namespace ExternalTypes
    {
        public class PersonInfo
        {
            public string FirstName { get; set; } = "";
            public string LastName { get; set; } = "";
            public int Age { get; set; }
        }

        public class CompanyInfo
        {
            public string Name { get; set; } = "";
            public string Location { get; set; } = "";
        }
    }

    // Root data model for the report.
    public class ReportModel
    {
        public List<ExternalTypes.PersonInfo> Persons { get; set; } = new();
        public ExternalTypes.CompanyInfo Company { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // 1. Create the template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Write static text and LINQ Reporting tags.
            builder.Writeln("Company: <<[model.Company.Name]>> located at <<[model.Company.Location]>>");
            builder.Writeln("Employees:");
            builder.Writeln("<<foreach [p in model.Persons]>>");
            builder.Writeln("- <<[p.FirstName]>> <<[p.LastName]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            string templatePath = Path.Combine(outputDir, "template.docx");
            template.Save(templatePath);

            // 2. Prepare the reporting engine and register external types.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(ExternalTypes.PersonInfo));
            engine.KnownTypes.Add(typeof(ExternalTypes.CompanyInfo));

            // Optional: enable reflection optimization (does not affect known type registration).
            ReportingEngine.UseReflectionOptimization = true;

            // 3. Prepare sample data.
            ReportModel model = new()
            {
                Company = new ExternalTypes.CompanyInfo
                {
                    Name = "Acme Corp",
                    Location = "New York"
                },
                Persons = new List<ExternalTypes.PersonInfo>
                {
                    new ExternalTypes.PersonInfo { FirstName = "John", LastName = "Doe", Age = 30 },
                    new ExternalTypes.PersonInfo { FirstName = "Jane", LastName = "Smith", Age = 28 }
                }
            };

            // 4. Load the template (demonstrating load before building).
            Document doc = new Document(templatePath);

            // 5. Build the report.
            engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            string outputPath = Path.Combine(outputDir, "output.docx");
            doc.Save(outputPath);

            // 7. Verify that the generated document contains expected values.
            string extractedText = doc.GetText();
            bool verificationPassed =
                extractedText.Contains("Acme Corp") &&
                extractedText.Contains("New York") &&
                extractedText.Contains("John Doe") &&
                extractedText.Contains("Jane Smith");

            Console.WriteLine(verificationPassed
                ? "Report generated successfully and verified."
                : "Verification failed: expected content not found.");
        }
    }
}
