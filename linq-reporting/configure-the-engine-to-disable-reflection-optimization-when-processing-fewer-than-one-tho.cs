using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingReflectionOptimization
{
    // Sample data model.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper model for the report.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for potential encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data with fewer than 1000 records.
            var model = new ReportModel();
            for (int i = 1; i <= 5; i++)
            {
                model.Persons.Add(new Person { Name = $"Person {i}", Age = 20 + i });
            }

            // Create a template document programmatically.
            var templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a simple foreach loop to list persons.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(templatePath);

            // Load the template for report generation.
            var reportDoc = new Document(templatePath);

            // Disable reflection optimization because we have fewer than 1000 records.
            ReportingEngine.UseReflectionOptimization = false;

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            var outputPath = "report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
