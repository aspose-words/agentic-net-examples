using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingMissingMembers
{
    // Simple data model with a defined property.
    public class Person
    {
        public string Name { get; set; } = string.Empty; // Initialized to avoid nullable warnings.
        // Note: No Age property – it will be missing in the template.
    }

    // Wrapper class that will be passed as the root data source.
    public class ReportModel
    {
        public Person Person { get; set; } = new(); // Initialized with target‑typed new.
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a template document programmatically.
            Document doc = new();
            DocumentBuilder builder = new(doc);

            // Insert LINQ Reporting tags. The Age tag does not exist on Person.
            builder.Writeln("Name: <<[model.Person.Name]>>");
            builder.Writeln("Age: <<[model.Person.Age]>>"); // Missing member.

            // 2. Prepare the data source.
            ReportModel model = new()
            {
                Person = new Person { Name = "John Doe" }
            };

            // 3. Configure the ReportingEngine to treat missing members as null.
            ReportingEngine engine = new()
            {
                Options = ReportBuildOptions.AllowMissingMembers
                // MissingMemberMessage can be left empty (default) to output nothing for missing members.
            };

            // 4. Build the report. The root name must match the tags ("model").
            engine.BuildReport(doc, model, "model");

            // 5. Save the generated report.
            const string outputPath = "ReportOutput.docx";
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
