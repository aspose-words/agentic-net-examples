using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

#nullable enable

namespace LinqReportingMissingItems
{
    // Data model classes
    public class Person
    {
        // Name may be null to simulate missing data
        public string? Name { get; set; }
    }

    public class Model
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // Create a template document with LINQ Reporting tags
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Persons list:");
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // Load the template for report generation
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // Prepare sample data with some missing (null) collection items
            // -----------------------------------------------------------------
            var model = new Model
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice" },
                    new Person { Name = null }, // Missing name should become empty string
                    new Person { Name = "Bob" }
                }
            };

            // -----------------------------------------------------------------
            // Configure the ReportingEngine (default behavior already treats null as empty)
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();

            // Build the report
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
