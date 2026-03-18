using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineRestrictedExample
{
    // Sample data class that contains a sensitive field.
    public class Employee
    {
        public string Name { get; set; }
        public decimal Salary { get; set; }   // Sensitive information.
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a simple template document with a placeholder for Salary.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            // The template tries to access the Salary property.
            builder.Writeln("Employee Name: <<[employee.Name]>>");
            builder.Writeln("Employee Salary: <<[employee.Salary]>>"); // Should be blocked.

            // 2. Restrict the Employee type so that its members cannot be accessed via the reporting engine.
            // This must be done before any report is built.
            ReportingEngine.SetRestrictedTypes(typeof(Employee));

            // 3. Prepare the data source.
            Employee employee = new Employee { Name = "John Doe", Salary = 12345.67m };

            // 4. Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members so that blocked members are rendered as empty strings
                // instead of throwing an exception.
                Options = ReportBuildOptions.AllowMissingMembers
            };
            // Optional: customize the message for missing members (empty string in this case).
            engine.MissingMemberMessage = string.Empty;

            // 5. Build the report.
            engine.BuildReport(doc, new { employee });

            // 6. Output the result to the console.
            Console.WriteLine("Generated Document Text:");
            Console.WriteLine(doc.GetText().Trim());

            // 7. Save the document if needed.
            doc.Save("RestrictedReport.docx");
        }
    }
}
