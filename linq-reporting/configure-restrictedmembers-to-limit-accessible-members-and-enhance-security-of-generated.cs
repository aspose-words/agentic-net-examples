using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used as the root object for the report.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some legacy encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Ensure output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Variable that attempts to obtain the base type of System.String.
            // This uses System.Type, which we will later restrict.
            builder.Writeln("<<var [typeVar = \"\".GetType().BaseType]>>");

            // Normal data fields.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Attempt to output the restricted type variable.
            builder.Writeln("Restricted type test: <<[typeVar]>>");

            // Save the template to disk.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Configure restricted members before building any report.
            // -----------------------------------------------------------------
            // Prevent access to members of System.Type (and its derived types) from the template.
            ReportingEngine.SetRestrictedTypes(typeof(System.Type));

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            Person person = new Person
            {
                Name = "John Doe",
                Age = 30
            };

            // -----------------------------------------------------------------
            // 4. Build the report.
            // -----------------------------------------------------------------
            // Load the template (demonstrates load‑save lifecycle).
            Document doc = new Document(templatePath);

            ReportingEngine engine = new ReportingEngine();
            // Allow missing members so that attempts to use restricted members
            // do not throw an exception but are treated as missing.
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            engine.MissingMemberMessage = "Restricted";

            // Build the report using the root object name "person".
            bool success = engine.BuildReport(doc, person, "person");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            string reportPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(reportPath);

            // Simple console output to indicate completion.
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report saved to: {reportPath}");
        }
    }
}
