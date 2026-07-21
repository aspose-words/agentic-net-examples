using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Root data model
    public class ReportModel
    {
        // Initialize collection to avoid null warnings
        public List<Person> Persons { get; set; } = new();
    }

    // Simple person entity
    public class Person
    {
        // Default to empty string to avoid nullable warnings
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -------------------------------------------------
            // Step 1: Create a template document with LINQ Reporting tags
            // -------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Report:");
            // Loop over the Persons collection
            builder.Writeln("<<foreach [person in Persons]>>");
            // Output only when the current item is not null
            builder.Writeln("<<if [person != null]>>Name: <<[person.Name]>> Age: <<[person.Age]>><</if>>");
            builder.Writeln("<</foreach>>");

            // Save the template locally
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // -------------------------------------------------
            // Step 2: Load the template for reporting
            // -------------------------------------------------
            var doc = new Document(templatePath);

            // -------------------------------------------------
            // Step 3: Prepare sample data with a null entry in the collection
            // -------------------------------------------------
            var model = new ReportModel();
            model.Persons.Add(new Person { Name = "Alice", Age = 30 });
            model.Persons.Add(null); // This null entry will be ignored by the IF condition
            model.Persons.Add(new Person { Name = "Bob", Age = 25 });

            // -------------------------------------------------
            // Step 4: Configure the ReportingEngine to treat missing members as empty strings
            // -------------------------------------------------
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            // Empty message suppresses output for missing members
            engine.MissingMemberMessage = string.Empty;

            // -------------------------------------------------
            // Step 5: Build the report
            // -------------------------------------------------
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // Step 6: Save the generated report
            // -------------------------------------------------
            const string outputPath = "output.docx";
            doc.Save(outputPath);

            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
