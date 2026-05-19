using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsRestrictedMembersDemo
{
    // Simple data model used in the report.
    public class Person
    {
        public string Name { get; set; }

        public Person(string name) => Name = name;

        // This member will be blocked by the restricted type configuration.
        public string GetSecret() => "TopSecret";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // The tags attempt to access members of the Person class.
            builder.Writeln("Name: <<[Name]>>");
            builder.Writeln("Secret: <<[GetSecret]>>");

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            Person person = new Person("John Doe");

            // -----------------------------------------------------------------
            // 3. Restrict the Person type so its members cannot be accessed.
            //    This must be done before any report is built.
            // -----------------------------------------------------------------
            ReportingEngine.SetRestrictedTypes(typeof(Person));

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine.
            //    - AllowMissingMembers prevents exceptions for blocked members.
            //    - MissingMemberMessage defines what is inserted for blocked members.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = "Restricted"
            };

            // -----------------------------------------------------------------
            // 5. Build the report.
            // -----------------------------------------------------------------
            engine.BuildReport(template, person);

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RestrictedReport.docx");
            template.Save(outputPath);

            // Optional: display the resulting plain text in the console.
            Console.WriteLine("Report generated. Extracted text:");
            Console.WriteLine(template.GetText().Trim());
        }
    }
}
