using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used in the template.
    public class Person
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a template document with a LINQ Reporting tag.
            string templatePath = "Template.docx";
            CreateTemplate(templatePath);

            // Step 2: Load the template document.
            Document template = new Document(templatePath);

            // Step 3: Configure restricted types to block potentially unsafe members.
            // This must be done before any report is built.
            ReportingEngine.SetRestrictedTypes(
                typeof(Process),
                typeof(Assembly),
                typeof(File),
                typeof(Directory),
                typeof(Environment));

            // Step 4: Prepare the data source.
            Person person = new Person { Name = "John Doe" };

            // Step 5: Build the report.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members to avoid exceptions if the template references something not present.
                Options = ReportBuildOptions.AllowMissingMembers
            };
            engine.BuildReport(template, person, "person");

            // Step 6: Save the generated report.
            string outputPath = "Report.docx";
            template.Save(outputPath);
        }

        // Creates a simple Word document containing a LINQ Reporting tag.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The tag references the 'Name' property of the root object named 'person'.
            builder.Writeln("Hello <<[person.Name]>>!");

            // Save the template to disk.
            doc.Save(filePath);
        }
    }
}
