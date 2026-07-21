using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with a sensitive Salary field.
    public class Employee
    {
        public string Name { get; set; } = "John Doe";
        public decimal Salary { get; set; } = 75000m; // Sensitive information.
        public string Position { get; set; } = "Software Engineer";
    }

    // Wrapper model used as the root object for the report.
    public class ReportModel
    {
        public Employee Employee { get; set; } = new Employee();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Employee Report");
            builder.Writeln("Name: <<[model.Employee.Name]>>");
            builder.Writeln("Position: <<[model.Employee.Position]>>");
            // Attempt to access the sensitive Salary field – this will be blocked.
            builder.Writeln("Salary: <<[model.Employee.Salary]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Configure restricted types before building the report.
            //    Restrict the Decimal type so that any decimal members (e.g., Salary)
            //    cannot be accessed from the template.
            // -----------------------------------------------------------------
            ReportingEngine.SetRestrictedTypes(typeof(decimal));

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var data = new ReportModel
            {
                Employee = new Employee
                {
                    Name = "Alice Smith",
                    Salary = 120000m, // This value should not appear in the output.
                    Position = "Project Manager"
                }
            };

            // -----------------------------------------------------------------
            // 4. Build the report.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine
            {
                // Allow missing members so that blocked members are rendered as empty strings.
                Options = ReportBuildOptions.AllowMissingMembers,
                // Optional: customize the message shown for blocked members.
                MissingMemberMessage = "[Restricted]"
            };

            // Load the template (could also reuse the in‑memory document, but the rule requires loading).
            var doc = new Document(templatePath);

            // Build the report using the root object name "model" as referenced in the template.
            engine.BuildReport(doc, data, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            // Output the plain text of the generated document to the console.
            Console.WriteLine("Generated report text:");
            Console.WriteLine(doc.GetText());
        }
    }
}
