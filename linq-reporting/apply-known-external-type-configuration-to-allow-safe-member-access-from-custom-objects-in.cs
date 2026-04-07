using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace AsposeWordsLinqReportingExample
{
    // Sample data model.
    public class Customer
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    // External static helper that will be referenced from the template.
    public static class ExternalHelper
    {
        public static bool IsAdult(int age) => age >= 18;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            string templatePath = "template.docx";
            string reportPath = "report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert simple text and LINQ Reporting tags.
            builder.Writeln("Customer Report");
            builder.Writeln("----------------");
            builder.Writeln("Name: <<[model.Name]>>");
            builder.Writeln("Age: <<[model.Age]>>");
            builder.Writeln("Adult? <<[ExternalHelper.IsAdult(model.Age)]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Register the external type so its static members can be used safely.
            engine.KnownTypes.Add(typeof(ExternalHelper));

            // Sample data source.
            Customer model = new Customer { Name = "John Doe", Age = 28 };

            // Build the report. The root object name must match the tag prefix used in the template.
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            loadedTemplate.Save(reportPath);
        }
    }
}
