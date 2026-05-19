using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model with a sensitive field.
    public class Employee
    {
        public Employee(string name, string position, decimal salary)
        {
            Name = name;
            Position = position;
            Salary = salary;
        }

        public string Name { get; set; }
        public string Position { get; set; }
        public decimal Salary { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Restrict the Employee type so its members cannot be accessed in templates.
            // -----------------------------------------------------------------
            ReportingEngine.SetRestrictedTypes(typeof(Employee));

            // -----------------------------------------------------------------
            // 2. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Employee Report");
            builder.Writeln("Name: <<[emp.Name]>>");
            builder.Writeln("Position: <<[emp.Position]>>");
            builder.Writeln("Salary: <<[emp.Salary]>>"); // This field will be restricted.

            // Save the template to disk.
            string templatePath = Path.Combine(outputDir, "template.docx");
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members so the engine does not throw when a restricted member is accessed.
                Options = ReportBuildOptions.AllowMissingMembers,
                // Message shown for restricted/missing members.
                MissingMemberMessage = "[Restricted]"
            };

            // Sample data.
            Employee emp = new Employee("John Doe", "Software Engineer", 95000m);

            // Build the report. The root object name used in the template is "emp".
            engine.BuildReport(report, emp, "emp");

            // Save the final document.
            string outputPath = Path.Combine(outputDir, "EmployeeReport.docx");
            report.Save(outputPath);
        }
    }
}
