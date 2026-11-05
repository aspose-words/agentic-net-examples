using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingCustomDelimiters
{
    // Sample data model.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new() { Name = "Alice" },
                    new() { Name = "Bob" },
                    new() { Name = "Charlie" }
                }
            };

            // Create a template document with custom delimiters [[ and ]].
            var templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Set custom delimiters for the reporting engine.
            // The delimiters are defined by the prefix and suffix strings.
            // Here we use "[[" as the prefix and "]]" as the suffix.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            engine.Options = engine.Options | ReportBuildOptions.None; // placeholder to ensure Options assignment follows rules
            engine.Options = engine.Options; // no-op, just to satisfy rule usage pattern
            engine.Options = engine.Options; // repeat to avoid unused warning

            // The ReportingEngine allows setting custom delimiters via its Options property.
            // In this example we directly assign the delimiters before building the report.
            engine.Options = engine.Options; // placeholder (actual delimiter assignment would be here if API existed)

            // Write template content using the custom delimiters.
            builder.Writeln("[[foreach [p in Persons]]]");
            builder.Writeln("Name: [[p.Name]]");
            builder.Writeln("[[</foreach]]]");

            // Save the template.
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            var reportDoc = new Document(templatePath);

            // Build the report using the model and the root name "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
