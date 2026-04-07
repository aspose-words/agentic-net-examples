using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used as the root object for the report.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a LINQ Reporting tag that attempts to expose a restricted type.
            // The tag creates a variable "typeVar" that holds the System.Diagnostics.Process type,
            // then writes the variable's value. Because Process is restricted, the engine will
            // block access and the output will be empty.
            builder.Writeln("Attempt to access restricted type:");
            builder.Writeln("<<var [typeVar = typeof(System.Diagnostics.Process)]>><<[typeVar]>>");

            // Insert a normal tag that displays data from the root object.
            builder.Writeln("Person name: <<[person.Name]>>");

            // Save the template to disk (required by the workflow).
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back for report generation.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            // Block access to types that could execute arbitrary code.
            ReportingEngine.SetRestrictedTypes(
                typeof(System.Diagnostics.Process),
                typeof(System.Reflection.Assembly),
                typeof(System.IO.File));

            // Create the engine and allow missing members so the report does not throw
            // when the restricted type is accessed.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers
            };
            engine.MissingMemberMessage = string.Empty; // No placeholder for blocked members.

            // -----------------------------------------------------------------
            // 4. Prepare the data source.
            // -----------------------------------------------------------------
            Person person = new Person { Name = "John Doe" };
            // The root object name used in the template is "person".
            // BuildReport will replace <<[person.Name]>> with the actual value.
            // The restricted type access will be silently ignored.
            engine.BuildReport(doc, person, "person");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "report.docx";
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
