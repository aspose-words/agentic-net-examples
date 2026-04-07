using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingErrorPlaceholders
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int? Age { get; set; } // Age may be missing (null) for demonstration.
    }

    // Wrapper class that will be passed as the root data source.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create the template document programmatically.
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            // Begin a foreach loop over the collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            // Normal fields.
            builder.Writeln("Name: <<[p.Name]>>");
            // Age field – may be missing (null) for some items.
            builder.Writeln("Age: <<[p.Age]>>");
            // Optional section that will be displayed only when Age has a value.
            builder.Writeln("<<if [p.Age != null]>>");
            builder.Writeln("  (Age is provided)");
            builder.Writeln("<</if>>");
            // End of the foreach loop.
            builder.Writeln("<</foreach>>");
            // Placeholder that will display any inline error messages produced by the engine.
            builder.Writeln("<<error>>");
            // Save the template to disk.
            builder.Document.Save(templatePath);

            // 2. Load the template back for reporting.
            var doc = new Document(templatePath);

            // 3. Prepare sample data – one person has Age, the other does not.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob" } // Age left null to simulate missing data.
                }
            };

            // 4. Configure the reporting engine to inline error messages.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };
            // Optional: customize the message shown for missing members.
            engine.MissingMemberMessage = "Missing";

            // 5. Build the report. The root object name is "model".
            bool success = engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            var outputPath = "ReportOutput.docx";
            doc.Save(outputPath);

            // Output simple console information (no user interaction required).
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
            Console.WriteLine($"Output saved to: {outputPath}");
        }
    }
}
