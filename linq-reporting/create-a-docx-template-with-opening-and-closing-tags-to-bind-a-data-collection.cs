using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model with a collection to bind.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the DOCX template programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            // Add a title.
            builder.Writeln("Persons Report");
            builder.Writeln();

            // Open a foreach block that iterates over the collection.
            builder.Writeln("<<foreach [person in Persons]>>");
            // Inside the block write the fields to be populated.
            builder.Writeln("- Name: <<[person.Name]>>, Age: <<[person.Age]>>");
            // Close the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            ReportModel model = new()
            {
                Persons = new List<Person>
                {
                    new() { Name = "Alice", Age = 30 },
                    new() { Name = "Bob", Age = 45 },
                    new() { Name = "Charlie", Age = 25 }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new(templatePath);

            ReportingEngine engine = new();
            // No special options are required for this simple example.
            engine.Options = ReportBuildOptions.None;

            // Build the report using the model as the root object named "model".
            bool success = engine.BuildReport(reportDoc, model, "model");

            // Optionally, you could check the success flag when using InlineErrorMessages.
            // For this example we simply proceed to save the document.

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
